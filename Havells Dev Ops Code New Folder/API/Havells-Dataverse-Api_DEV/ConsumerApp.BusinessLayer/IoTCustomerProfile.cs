using System;
using Microsoft.Xrm.Sdk;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ConsumerApp.BusinessLayer
{
    public class IoTCustomerProfile
    {
        [DataMember]
        public Guid CustomerGuid { get; set; }
        [DataMember]
        public string MobileNumber { get; set; }
        [DataMember]
        public string SalutationName { get; set; }
        [DataMember]
        public int? Salutation { get; set; }
        [DataMember]
        public string FirstName { get; set; }
        [DataMember]
        public string MiddleName { get; set; }
        [DataMember]
        public string LastName { get; set; }
        [DataMember]
        public bool? Gender { get; set; }
        [DataMember]
        public string GenderName { get; set; }
        [DataMember]
        public string DateOfBirth { get; set; }
        [DataMember]
        public string DateOfAnniversary { get; set; }
        [DataMember]
        public string Email { get; set; }
        [DataMember]
        public string AlternateNumber { get; set; }
        [DataMember]
        public bool? Consent { get; set; }
        [DataMember]
        public bool? SubscribeForMsgService { get; set; } // hil_subscribeformessagingservice
        [DataMember]
        public string PreferredLanguage { get; set; } // hil_preferredlanguageforcommunication
        [DataMember]
        public int? CustomerType { get; set; } // Customer Type (1-Consumer,2-Dealer)
        [DataMember]
        public bool? IsPremiumCustomer { get; set; } // (true-Yes,false-No)
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }

        public IoTCustomerProfile GetIoTCustomerProfile(IoTCustomerProfile customerProfileData)
        {
            IoTCustomerProfile customerProfile = new IoTCustomerProfile();
            EntityCollection entcoll;
            QueryExpression Query;
            Regex Regex_MobileNo = new Regex("^[6-9]\\d{9}$");

            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (string.IsNullOrWhiteSpace(customerProfileData.CustomerGuid.ToString()) || customerProfileData.CustomerGuid == Guid.Empty || customerProfileData.CustomerGuid.ToString().Trim().Length == 0)
                    {
                        if (string.IsNullOrWhiteSpace(customerProfileData.MobileNumber))
                        {
                            return new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Customer Mobile Number is required." };
                        }
                        else if (!Regex_MobileNo.IsMatch(customerProfileData.MobileNumber))
                        {
                            return new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Invalid Customer Mobile Number." };
                        }
                        else
                        {
                            Query = new QueryExpression("contact");
                            Query.ColumnSet = new ColumnSet(false);
                            Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, customerProfileData.MobileNumber);
                            entcoll = service.RetrieveMultiple(Query);
                            if (entcoll.Entities.Count > 0)
                            {
                                customerProfileData.CustomerGuid = entcoll.Entities[0].Id;
                            }
                            else
                            {
                                return new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Customer does not exist." };
                            }
                        }
                    }
                    Entity ent = null;
                    LinkEntity _lnkLang = new LinkEntity()
                    {
                        Columns = new ColumnSet("hil_code"),
                        EntityAlias = "lang",
                        LinkFromEntityName = "contact",
                        LinkFromAttributeName = "hil_preferredlanguageforcommunication",
                        LinkToEntityName = "hil_preferredlanguageforcommunication",
                        LinkToAttributeName = "hil_preferredlanguageforcommunicationid",
                        JoinOperator = JoinOperator.LeftOuter
                    };

                    Query = new QueryExpression("contact");
                    Query.ColumnSet = new ColumnSet("mobilephone", "hil_salutation", "firstname", "middlename", "lastname", "hil_gender", "hil_dateofbirth", "hil_dateofanniversary", "emailaddress1", "address1_telephone3", "hil_consent", "hil_subscribeformessagingservice", "hil_preferredlanguageforcommunication", "hil_premiumcustomer");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("contactid", ConditionOperator.Equal, customerProfileData.CustomerGuid);
                    Query.LinkEntities.Add(_lnkLang);

                    entcoll = service.RetrieveMultiple(Query);

                    if (entcoll.Entities.Count == 0)
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Customer does not exist." };
                        return customerProfile;
                    }
                    else
                    {
                        ent = entcoll.Entities[0];
                        customerProfile.CustomerGuid = customerProfileData.CustomerGuid;
                        if (ent.Attributes.Contains("mobilephone"))
                            customerProfile.MobileNumber = ent.GetAttributeValue<string>("mobilephone");
                        if (ent.Attributes.Contains("firstname"))
                            customerProfile.FirstName = ent.GetAttributeValue<string>("firstname");
                        if (ent.Attributes.Contains("middlename"))
                            customerProfile.MiddleName = ent.GetAttributeValue<string>("middlename");
                        if (ent.Attributes.Contains("lastname"))
                            customerProfile.LastName = ent.GetAttributeValue<string>("lastname");
                        if (ent.Attributes.Contains("hil_gender"))
                        {
                            bool gender = ent.GetAttributeValue<bool>("hil_gender");
                            customerProfile.Gender = gender;
                            if (gender) { customerProfile.GenderName = "Female"; } //True
                            else { customerProfile.GenderName = "Male"; } // False
                        }
                        if (ent.Attributes.Contains("hil_dob"))
                            customerProfile.DateOfBirth = ent.GetAttributeValue<DateTime>("hil_dob").AddDays(1).ToString("mm-dd-yyyy");
                        if (ent.Attributes.Contains("hil_doa"))
                            customerProfile.DateOfAnniversary = ent.GetAttributeValue<DateTime>("hil_doa").AddDays(1).ToString("mm-dd-yyyy");
                        if (ent.Attributes.Contains("emailaddress1"))
                            customerProfile.Email = ent.GetAttributeValue<string>("emailaddress1");
                        if (ent.Attributes.Contains("address1_telephone3"))
                            customerProfile.AlternateNumber = ent.GetAttributeValue<string>("address1_telephone3");
                        if (ent.Attributes.Contains("hil_salutation"))
                        {
                            customerProfile.Salutation = ent.GetAttributeValue<OptionSetValue>("hil_salutation").Value;
                            customerProfile.SalutationName = ent.FormattedValues["hil_salutation"].ToString();
                        }
                        if (ent.Attributes.Contains("hil_consent"))
                            customerProfile.Consent = ent.GetAttributeValue<bool>("hil_consent");
                        else
                            customerProfile.Consent = false;

                        if (ent.Attributes.Contains("hil_subscribeformessagingservice"))
                            customerProfile.SubscribeForMsgService = ent.GetAttributeValue<bool>("hil_subscribeformessagingservice");
                        else
                            customerProfile.SubscribeForMsgService = false;

                        if (ent.Attributes.Contains("hil_preferredlanguageforcommunication"))
                            customerProfile.PreferredLanguage = ent.GetAttributeValue<AliasedValue>("lang.hil_code").Value.ToString().Trim();
                        //if (ent.Attributes.Contains("hil_customertype"))
                        //{
                        //    customerProfile.CustomerType = ent.GetAttributeValue<OptionSetValue>("hil_customertype").Value;
                        //}
                        if (ent.Attributes.Contains("hil_premiumcustomer"))
                        {
                            customerProfile.IsPremiumCustomer = ent.GetAttributeValue<bool>("hil_premiumcustomer");
                        }
                        customerProfile.StatusCode = "200";
                        customerProfile.StatusDescription = "OK";
                    }
                }
                else
                {
                    customerProfile = new IoTCustomerProfile { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                customerProfile = new IoTCustomerProfile { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return customerProfile;
        }

        public IoTCustomerProfile UpdateIoTCustomerProfile(IoTCustomerProfile customerProfileData)
        {
            IoTCustomerProfile customerProfile;
            //EntityCollection entcoll;
            //QueryExpression Query;

            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                EntityReference _entPreferredLang = null;
                if (service != null)
                {
                    if (customerProfileData.CustomerGuid.ToString().Trim().Length == 0)
                    {
                        return new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Customer GUID is required." };
                    }
                    if (customerProfileData.CustomerType != null)
                    {
                        if (customerProfileData.CustomerType < 1 || customerProfileData.CustomerType > 2)
                        {
                            return new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Invalid Consumer Type value. Possible values would be (1-Consumer,2-Dealer)" };
                        }
                    }
                    if (customerProfileData.PreferredLanguage != null)
                    {
                        QueryExpression qspreferLang = new QueryExpression("hil_preferredlanguageforcommunication");
                        qspreferLang.ColumnSet = new ColumnSet(false);
                        ConditionExpression condExp = new ConditionExpression("hil_code", ConditionOperator.Equal, customerProfileData.PreferredLanguage);
                        qspreferLang.Criteria.AddCondition(condExp);
                        EntityCollection entColPreferLang = service.RetrieveMultiple(qspreferLang);
                        if (entColPreferLang.Entities.Count == 0)
                        {
                            return new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Preferred Language does not exist." };
                        }
                        else
                        {
                            _entPreferredLang = entColPreferLang.Entities[0].ToEntityReference();
                        }
                    }
                    Entity ent = service.Retrieve("contact", customerProfileData.CustomerGuid, new ColumnSet("mobilephone"));
                    if (ent == null)
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Customer does not exist." };
                        return customerProfile;
                    }
                    else
                    {
                        if (customerProfileData.MobileNumber != null && customerProfileData.MobileNumber.ToString().Trim().Length > 0)
                        {
                            if (customerProfileData.MobileNumber != ent.GetAttributeValue<string>("mobilephone"))
                            {
                                customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Mobile Number does not belong to Customer." };
                                return customerProfile;
                            }
                        }
                        Contact entCustomer = new Contact();
                        entCustomer.Id = ent.Id;
                        if (customerProfileData.Salutation != null)
                        {
                            List<HashTableDTO> salutationList = new List<HashTableDTO>();
                            IoTCommonLib commonLib = new IoTCommonLib();
                            if (!commonLib.GetSalutationEnum().Exists(x => x.Value == customerProfileData.Salutation))
                            {
                                customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Salutation not found in D365. To get list of valid Salutations,Please Call <GetSalutationEnum> API and pass proper Salutation." };
                                return customerProfile;
                            }
                            entCustomer.hil_Salutation = new OptionSetValue(customerProfileData.Salutation.Value);
                        }
                        if (customerProfileData.FirstName != null)
                            entCustomer.FirstName = customerProfileData.FirstName;
                        if (customerProfileData.MiddleName != null)
                            entCustomer.MiddleName = customerProfileData.MiddleName;
                        if (customerProfileData.LastName != null)
                            entCustomer.LastName = customerProfileData.LastName;
                        if (customerProfileData.AlternateNumber != null)
                            entCustomer["address1_telephone3"] = customerProfileData.AlternateNumber;
                        if (customerProfileData.Gender != null)
                            entCustomer.hil_Gender = !customerProfileData.Gender; //{True:Male,False:Female}

                        if (customerProfileData.Email != null)
                            entCustomer.EMailAddress1 = customerProfileData.Email;

                        if (customerProfileData.DateOfBirth != null)
                        {
                            DateTime dtDOB = DateTime.ParseExact(customerProfileData.DateOfBirth, new string[] { "MM.dd.yyyy", "MM-dd-yyyy", "MM/dd/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None);
                            entCustomer.hil_DOB = dtDOB;
                        }
                        if (customerProfileData.DateOfAnniversary != null)
                        {
                            DateTime dtDOA = DateTime.ParseExact(customerProfileData.DateOfAnniversary, new string[] { "MM.dd.yyyy", "MM-dd-yyyy", "MM/dd/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None);
                            entCustomer.hil_DOA = dtDOA;
                        }
                        if (customerProfileData.Consent != null)
                            entCustomer["hil_consent"] = customerProfileData.Consent;

                        if (customerProfileData.SubscribeForMsgService != null)
                            entCustomer["hil_subscribeformessagingservice"] = customerProfileData.SubscribeForMsgService;

                        if (_entPreferredLang != null)
                            entCustomer["hil_preferredlanguageforcommunication"] = _entPreferredLang;
                        //if (customerProfileData.CustomerType != null)
                        //    entCustomer["hil_customertype"] = new OptionSetValue(Convert.ToInt16(customerProfileData.CustomerType));
                        if (customerProfileData.IsPremiumCustomer != null)
                            entCustomer["hil_premiumcustomer"] = customerProfileData.IsPremiumCustomer;

                        service.Update(entCustomer);
                        customerProfile = new IoTCustomerProfile { MobileNumber = ent.GetAttributeValue<string>("mobilephone"), CustomerGuid = ent.Id, StatusCode = "200", StatusDescription = "OK." };
                    }
                }
                else
                {
                    customerProfile = new IoTCustomerProfile { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                customerProfile = new IoTCustomerProfile { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return customerProfile;
        }
    }
}
