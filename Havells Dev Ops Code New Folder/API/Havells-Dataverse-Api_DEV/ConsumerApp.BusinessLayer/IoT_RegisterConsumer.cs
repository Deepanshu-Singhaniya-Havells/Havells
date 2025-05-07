using System;
using Microsoft.Xrm.Sdk;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class IoT_RegisterConsumer
    {
        [DataMember]
        public string MobileNumber { get; set; }

        [DataMember]
        public string Email { get; set; }

        [DataMember]
        public string FirstName { get; set; }

        [DataMember]
        public int? Salutation { get; set; }

        [DataMember]
        public int? SourceOfCreation { get; set; } //4 for Whatsapp 5 for Iot Platform 7 for eCommerce 8 for ChatBot, 9 for Experience Store, 10 for Dealer Portal, 17 for Sampark
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

        public ReturnResult RegisterConsumer(IoT_RegisterConsumer consumer)
        {
            ReturnResult returnResult = new ReturnResult();
            Guid consumerGuId = Guid.Empty;
            EntityReference _entPreferredLang = null;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService(); //get org service obj for connection
                if (string.IsNullOrWhiteSpace(consumer.MobileNumber))
                {
                    returnResult.StatusCode = "204";
                    returnResult.StatusDescription = "No Content : Mobile Number is required.";
                    return returnResult;
                }
                if (consumer.PreferredLanguage != null)
                {
                    QueryExpression qspreferLang = new QueryExpression("hil_preferredlanguageforcommunication");
                    qspreferLang.ColumnSet = new ColumnSet(false);
                    ConditionExpression condExp = new ConditionExpression("hil_code", ConditionOperator.Equal, consumer.PreferredLanguage.Trim());
                    qspreferLang.Criteria.AddCondition(condExp);
                    EntityCollection entColPreferLang = service.RetrieveMultiple(qspreferLang);
                    if (entColPreferLang.Entities.Count == 0)
                    {
                        returnResult.StatusCode = "204";
                        returnResult.StatusDescription = "No Content : Preferred Language does not exist.";
                        return returnResult;
                    }
                    else
                    {
                        _entPreferredLang = entColPreferLang.Entities[0].ToEntityReference();
                    }
                }
                if (consumer.CustomerType != null)
                {
                    if (consumer.CustomerType < 1 || consumer.CustomerType > 2)
                    {
                        returnResult.StatusCode = "204";
                        returnResult.StatusDescription = "Invalid Consumer Type value. Possible values would be (1-Consumer,2-Dealer)";
                        return returnResult;
                    }
                }
                if (consumer.SourceOfCreation == null)
                {
                    returnResult.StatusCode = "204";
                    returnResult.StatusDescription = "No Content : Source of Registration is required. Please pass <4> for Whatsapp <5> for IoT Platform <7> for eCommerce<8> for Chatbot";
                    return returnResult;
                }
                else
                {
                    if (service != null)
                    {
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

                        //Checking Mobile Number already exist
                        QueryExpression qsContact = new QueryExpression("contact");
                        qsContact.ColumnSet = new ColumnSet("contactid", "fullname", "emailaddress1", "hil_salutation", "hil_consent", "hil_subscribeformessagingservice", "hil_preferredlanguageforcommunication", "hil_premiumcustomer");
                        ConditionExpression condExp = new ConditionExpression("mobilephone", ConditionOperator.Equal, consumer.MobileNumber);
                        qsContact.Criteria.AddCondition(condExp);
                        qsContact.LinkEntities.Add(_lnkLang);

                        EntityCollection entColConsumer = service.RetrieveMultiple(qsContact);
                        if (entColConsumer.Entities.Count > 0) //Consumer already Exists in D365 Database
                        {
                            consumerGuId = entColConsumer.Entities[0].Id;
                            returnResult.MobileNumber = consumer.MobileNumber;
                            returnResult.CustomerName = entColConsumer.Entities[0].GetAttributeValue<string>("fullname");
                            returnResult.EmailId = entColConsumer.Entities[0].GetAttributeValue<string>("emailaddress1");
                            if (entColConsumer.Entities[0].Attributes.Contains("hil_consent"))
                                returnResult.Consent = entColConsumer.Entities[0].GetAttributeValue<bool>("hil_consent");
                            else
                                returnResult.Consent = false;

                            if (entColConsumer.Entities[0].Attributes.Contains("hil_subscribeformessagingservice"))
                                returnResult.SubscribeForMsgService = entColConsumer.Entities[0].GetAttributeValue<bool>("hil_subscribeformessagingservice");
                            else
                                returnResult.SubscribeForMsgService = false;
                            if (entColConsumer.Entities[0].Attributes.Contains("hil_preferredlanguageforcommunication"))
                            {
                                returnResult.PreferredLanguage = entColConsumer.Entities[0].GetAttributeValue<AliasedValue>("lang.hil_code").Value.ToString().Trim();
                            }
                            //if (entColConsumer.Entities[0].Attributes.Contains("hil_customertype"))
                            //{
                            //    returnResult.CustomerType = entColConsumer.Entities[0].GetAttributeValue<OptionSetValue>("hil_customertype").Value;
                            //}
                            if (entColConsumer.Entities[0].Attributes.Contains("hil_premiumcustomer"))
                            {
                                returnResult.IsPremiumCustomer = entColConsumer.Entities[0].GetAttributeValue<bool>("hil_premiumcustomer");
                            }
                            if (consumer.SourceOfCreation == 11) // Voice Bot - 11
                            {
                                //Getting Last Served Address of the Consumer
                                QueryExpression qsAddress = new QueryExpression("hil_address");
                                qsAddress.ColumnSet = new ColumnSet("hil_fulladdress", "hil_pincode");
                                qsAddress.Criteria.AddCondition(new ConditionExpression("hil_customer", ConditionOperator.Equal, entColConsumer.Entities[0].Id));
                                qsAddress.TopCount = 1;
                                qsAddress.AddOrder("modifiedon", OrderType.Descending);
                                EntityCollection entColAddress = service.RetrieveMultiple(qsAddress);
                                if (entColAddress.Entities.Count > 0)
                                {
                                    returnResult.PINCode = entColAddress.Entities[0].GetAttributeValue<EntityReference>("hil_pincode").Name;
                                    returnResult.Address = entColAddress.Entities[0].GetAttributeValue<string>("hil_fulladdress");
                                }
                                else // If Address does not exist then send PINCode as NA 
                                {
                                    returnResult.PINCode = "NA";
                                }
                            }
                            returnResult.StatusCode = "208";
                            returnResult.StatusDescription = "Already Reported";
                            returnResult.CustomerGuid = consumerGuId;
                        }
                        if (consumerGuId == Guid.Empty) //Creating Consumer in D365 Database
                        {
                            if ((consumer.SourceOfCreation == 4 || consumer.SourceOfCreation == 7 || consumer.SourceOfCreation == 8 || consumer.SourceOfCreation == 9 
                                || consumer.SourceOfCreation == 10 || consumer.SourceOfCreation == 11 || consumer.SourceOfCreation == 17) 
                                && (consumer.FirstName == null || consumer.FirstName.Trim().Length == 0)) //If Source Of creation is 4,7 or 8,9,10 then check for First Name
                            {
                                returnResult.StatusCode = "204";
                                returnResult.StatusDescription = "No Content : Customer Name is required.";
                            }
                            else
                            {
                                Entity entConsumer = new Entity("contact");
                                entConsumer["mobilephone"] = consumer.MobileNumber;

                                if (consumer.Salutation != null)
                                {
                                    List<HashTableDTO> salutationList = new List<HashTableDTO>();
                                    IoTCommonLib commonLib = new IoTCommonLib();
                                    salutationList = commonLib.GetSalutationEnum();
                                    if (!salutationList.Exists(x => x.Value == consumer.Salutation))
                                    {
                                        returnResult.StatusCode = "204";
                                        returnResult.StatusDescription = "Salutation not found in D365. Please Call <GetSalutationEnum> API and pass proper Salutation.";
                                        return returnResult;
                                    }
                                    entConsumer["hil_salutation"] = new OptionSetValue(consumer.Salutation.Value);
                                }

                                if (string.IsNullOrWhiteSpace(consumer.FirstName))
                                {
                                    entConsumer["firstname"] = "UNDEF-" + consumer.MobileNumber;
                                }
                                else
                                {
                                    string[] consumerName = consumer.FirstName.Split(' ');
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
                                        entConsumer["firstname"] = consumer.FirstName;
                                    }
                                }

                                if (consumer.Email != null && consumer.Email.Trim().Length > 0)
                                {
                                    entConsumer["emailaddress1"] = consumer.Email;
                                }

                                entConsumer["hil_consumersource"] = new OptionSetValue(SourceOfCreation.Value);

                                if (consumer.Consent != null)
                                    entConsumer["hil_consent"] = consumer.Consent;

                                if (consumer.SubscribeForMsgService != null)
                                    entConsumer["hil_subscribeformessagingservice"] = consumer.SubscribeForMsgService;

                                if (_entPreferredLang != null)
                                    entConsumer["hil_preferredlanguageforcommunication"] = _entPreferredLang;
                                //if (consumer.CustomerType != null)
                                //    entConsumer["hil_customertype"] = new OptionSetValue(Convert.ToInt16(consumer.CustomerType));
                                if (consumer.IsPremiumCustomer != null)
                                    entConsumer["hil_premiumcustomer"] = consumer.IsPremiumCustomer;

                                consumerGuId = service.Create(entConsumer);
                                returnResult.CustomerGuid = consumerGuId;
                                returnResult.StatusCode = "200";
                                returnResult.StatusDescription = "OK";
                            }
                        }
                    }
                    else
                    {
                        returnResult.StatusCode = "503";
                        returnResult.StatusDescription = "D365 Service Unavailable";
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

        public PreferredLanguage GetPreferredLanguages()
        {
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                EntityCollection entcoll;

                QueryExpression query = new QueryExpression("hil_preferredlanguageforcommunication");
                query.ColumnSet = new ColumnSet("hil_code", "hil_name");
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                query.AddOrder("hil_name", OrderType.Ascending);
                entcoll = service.RetrieveMultiple(query);
                if (entcoll.Entities.Count == 0)
                {
                    return new PreferredLanguage { StatusCode = "204", StatusDescription = "Preferred Language found." };
                }
                else
                {
                    List<PreferredLanguageResponse> lstResponse = new List<PreferredLanguageResponse>();
                    foreach (Entity ent in entcoll.Entities)
                    {
                        if (ent.Attributes.Contains("hil_name") && ent.Attributes.Contains("hil_code"))
                        {
                            lstResponse.Add(
                            new PreferredLanguageResponse()
                            {
                                LangCode = ent.GetAttributeValue<string>("hil_code"),
                                LangName = ent.GetAttributeValue<string>("hil_code")
                            });
                        }
                    }
                    return new PreferredLanguage { PreferredLanguages = lstResponse, StatusCode = "200", StatusDescription = "Success" };
                }
            }
            catch (Exception ex)
            {
                return new PreferredLanguage { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message };
            }
        }
    }
    [DataContract]
    public class CustomerMaster
    {
        [DataMember]
        public string MobileNumber { get; set; }
        [DataMember]
        public string Email { get; set; }
        [DataMember]
        public string FullName { get; set; }
        [DataMember]
        public bool? Gender { get; set; }
        [DataMember]
        public string DateOfBirth { get; set; }
        [DataMember]
        public string DateOfAnniversary { get; set; }
        [DataMember]
        public string AlternateNumber { get; set; }
        [DataMember]
        public int? SourceOfCreation { get; set; }
        [DataMember]
        public bool? Consent { get; set; }
        [DataMember]
        public bool? SubscribeForMsgService { get; set; } // hil_subscribeformessagingservice
        [DataMember]
        public string PreferredLanguage { get; set; } // hil_preferredlanguageforcommunication
        [DataMember]
        public string AddressLine1 { get; set; }
        [DataMember]
        public string AddressLine2 { get; set; }
        [DataMember]
        public string AddressLine3 { get; set; }
        [DataMember]
        public string AddressPhone { get; set; }
        [DataMember]
        public string PINCode { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }

    [DataContract]
    public class PreferredLanguageResponse
    {
        [DataMember]
        public string LangCode { get; set; }
        [DataMember]
        public string LangName { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }

    [DataContract]
    public class PreferredLanguage
    {
        [DataMember]
        public List<PreferredLanguageResponse> PreferredLanguages { get; set; }

        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }
}
