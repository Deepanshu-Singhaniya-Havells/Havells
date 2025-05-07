using Havells.Dataverse.CustomConnector.Customer_Asset;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Consumer_Profile
{
    public class IotRegisterConsumer : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            string JsonResponse = "";
            string MobileNumber = string.Empty;
            string Email = string.Empty;
            string FirstName = string.Empty;
            int Salutation = 0;
            int SourceOfCreation = 0;
            string PreferredLanguage = null;
            bool isValidSourceOfCreation = false;
            try
            {
                if (context.InputParameters.Contains("MobileNumber") && context.InputParameters["MobileNumber"] is string)
                {

                    MobileNumber = Convert.ToString(context.InputParameters["MobileNumber"]);
                    if (string.IsNullOrWhiteSpace(MobileNumber))
                    {
                        string msg = "Mobile Number is required.";
                        Response(msg, context);
                        return;
                    }
                    else if (!APValidate.IsValidMobileNumber(MobileNumber))
                    {
                        string msg = "Invalid Mobile Number.";
                        Response(msg, context);
                        return;
                    }
                    if (context.InputParameters.Contains("SourceOfCreation"))
                    {
                        isValidSourceOfCreation = int.TryParse(context.InputParameters["SourceOfCreation"].ToString(), out SourceOfCreation);
                        if (isValidSourceOfCreation)
                            SourceOfCreation = Convert.ToInt32(context.InputParameters["SourceOfCreation"].ToString());
                    }
                    if (!isValidSourceOfCreation || SourceOfCreation == 0)
                    {
                        Response("Source of Registration is required. Please pass <4> for Whatsapp <5> for IoT Platform <7> for eCommerce<8> for Chatbot", context);
                        return;
                    }
                    FirstName = Convert.ToString(context.InputParameters["FirstName"]);
                    if (!APValidate.IsValidString(FirstName) && !string.IsNullOrWhiteSpace(FirstName))
                    {
                        Response("Invalid First Name", context);
                        return;
                    }
                    Salutation = Convert.ToInt32(context.InputParameters["Salutation"]);
                    {
                        if (!OptionSetValues.SalutationItem.Contains(Salutation) && Salutation != 0)
                        {
                            _tracingService.Trace("Salution Value " + Salutation);
                            Response("Invalid Salutation Value", context);
                            return;
                        }
                        bool Consent = Convert.ToBoolean(context.InputParameters["Consent"]);
                        bool SubscribeForMsgService = Convert.ToBoolean(context.InputParameters["SubscribeForMsgService"]);
                        Email = Convert.ToString(context.InputParameters["Email"]);
                        _tracingService.Trace("Email Value " + Email);
                        {
                            if (!APValidate.IsValidEmail(Email) && !string.IsNullOrWhiteSpace(Email))

                            {
                                Response("Invalid Email Address", context);
                                return;
                            }

                            if (context.InputParameters.Contains("PreferredLanguage"))
                            {
                                if (!string.IsNullOrWhiteSpace(Convert.ToString(context.InputParameters["PreferredLanguage"])))
                                    PreferredLanguage = Convert.ToString(context.InputParameters["PreferredLanguage"]);
                            }
                            JsonResponse = JsonSerializer.Serialize(RegisterConsumer(MobileNumber, Email, FirstName, Salutation, SourceOfCreation, Consent, SubscribeForMsgService, PreferredLanguage, service));
                            _tracingService.Trace(JsonResponse);
                            context.OutputParameters["data"] = JsonResponse;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var retObj = JsonSerializer.Serialize(new IoTValidateSerialNumber { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() });
                context.OutputParameters["data"] = retObj;
                return;
            }
        }
        public static void Response(string msg, IPluginExecutionContext context)
        {
            string JsonResponse = "";
            JsonResponse = JsonSerializer.Serialize(new ReturnResult
            {
                StatusCode = "204",
                StatusDescription = msg
            });
            context.OutputParameters["data"] = JsonResponse;
            return;
        }
        public static ReturnResult RegisterConsumer(string MobileNumber, string Email, string FirstName, int Salutation, int SourceOfCreation, bool? Consent, bool? SubscribeForMsgService, string PreferredLanguage, IOrganizationService service)
        {
            ReturnResult returnResult = new ReturnResult();
            Guid consumerGuId = Guid.Empty;
            EntityReference _entPreferredLang = null;
            try
            {
                if (PreferredLanguage != null)
                {
                    QueryExpression qspreferLang = new QueryExpression("hil_preferredlanguageforcommunication");
                    qspreferLang.ColumnSet = new ColumnSet(false);
                    ConditionExpression condExp = new ConditionExpression("hil_code", ConditionOperator.Equal, PreferredLanguage.Trim());
                    qspreferLang.Criteria.AddCondition(condExp);
                    EntityCollection entColPreferLang = service.RetrieveMultiple(qspreferLang);
                    if (entColPreferLang.Entities.Count == 0)
                    {
                        returnResult.StatusCode = "204";
                        returnResult.StatusDescription = " Preferred Language does not exist.";
                        return returnResult;
                    }
                    else
                    {
                        _entPreferredLang = entColPreferLang.Entities[0].ToEntityReference();
                    }
                }
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
                    qsContact.ColumnSet = new ColumnSet("contactid", "fullname", "emailaddress1", "hil_salutation", "hil_consent", "hil_subscribeformessagingservice", "hil_preferredlanguageforcommunication");
                    FilterExpression filterExpression = new FilterExpression(LogicalOperator.Or);
                    filterExpression.AddCondition(new ConditionExpression("mobilephone", ConditionOperator.Equal, MobileNumber));
                    filterExpression.AddCondition(new ConditionExpression("emailaddress1", ConditionOperator.Equal, Email));
                    qsContact.Criteria = filterExpression;
                    qsContact.LinkEntities.Add(_lnkLang);

                    EntityCollection entColConsumer = service.RetrieveMultiple(qsContact);
                    if (entColConsumer.Entities.Count > 0) //Consumer already Exists in D365 Database
                    {
                        consumerGuId = entColConsumer.Entities[0].Id;
                        returnResult.MobileNumber = MobileNumber;
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
                        if (SourceOfCreation == 11) // Voice Bot - 11
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
                        if ((SourceOfCreation == 4 || SourceOfCreation == 7 || SourceOfCreation == 8 || SourceOfCreation == 9 || SourceOfCreation == 10 || SourceOfCreation == 11 || SourceOfCreation == 17)
                            && (FirstName == null || FirstName.Trim().Length == 0)) //If Source Of creation is 4,7 or 8,10 then check for First Name
                        {
                            returnResult.StatusCode = "204";
                            returnResult.StatusDescription = "Customer Name is required.";
                        }
                        else
                        {
                            Entity entConsumer = new Entity("contact");
                            entConsumer["mobilephone"] = MobileNumber;

                            if (Salutation != 0)
                            {
                                List<HashTableDTO> salutationList = new List<HashTableDTO>();
                                IoTCommonLib commonLib = new IoTCommonLib();
                                salutationList = commonLib.GetSalutationEnum(service);
                                if (!salutationList.Exists(x => x.Value == Salutation))
                                {
                                    returnResult.StatusCode = "204";
                                    returnResult.StatusDescription = "Salutation not found in D365. Please Call <GetSalutationEnum> API and pass proper Salutation.";
                                    return returnResult;
                                }
                                entConsumer["hil_salutation"] = new OptionSetValue(Salutation);
                            }

                            if (FirstName == null)
                            {
                                entConsumer["firstname"] = "UNDEF-IoT-" + MobileNumber;
                            }
                            else
                            {
                                string[] consumerName = FirstName.Split(' ');
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
                                    entConsumer["firstname"] = FirstName;
                                }
                            }

                            if (Email != null && Email.Trim().Length > 0)
                            {
                                entConsumer["emailaddress1"] = Email;
                            }

                            entConsumer["hil_consumersource"] = new OptionSetValue(SourceOfCreation);
                            if (Consent != null)
                                entConsumer["hil_consent"] = Consent;

                            if (SubscribeForMsgService != null)
                                entConsumer["hil_subscribeformessagingservice"] = SubscribeForMsgService;

                            if (_entPreferredLang != null)
                                entConsumer["hil_preferredlanguageforcommunication"] = _entPreferredLang;

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
            catch (Exception ex)
            {
                returnResult.StatusCode = "500";
                returnResult.StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper();
            }
            return returnResult;
        }
    }
    public class ReturnResult
    {
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
        public Guid CustomerGuid { get; set; }
        public string MobileNumber { get; set; }
        public string CustomerName { get; set; }
        public string EmailId { get; set; }
        public bool? Consent { get; set; }
        public bool? SubscribeForMsgService { get; set; } // hil_subscribeformessagingservice    
        public string PreferredLanguage { get; set; } // hil_preferredlanguageforcommunication
        public string PINCode { get; set; }
        public string Address { get; set; }
    }
    public class HashTableDTO
    {
        public string Label { get; set; }
        public int? Value { get; set; }
        public string Extension { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
    }
    public class HashTableGuidDTO
    {
        public string Label { get; set; }
        public Guid Value { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
    }
    public class ProductHierarchyDTO
    {
        public string ProductCategory { get; set; }
        public Guid ProductCategoryGuid { get; set; }
        public string ProductSubCategory { get; set; }
        public Guid ProductSubCategoryGuid { get; set; }
        public bool IsSerialized { get; set; }
        public bool IsVerificationrequired { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
    }
    public class ProductDTO
    {
        public string ProductCode { get; set; }
        public string Product { get; set; }
        public Guid ProductGuid { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
    }
    public class IoTCommonLib
    {
        public List<HashTableDTO> GetSalutationEnum(IOrganizationService service)
        {
            HashTableDTO objIoTSalutationEnum;
            List<HashTableDTO> lstIoTSalutationEnum = new List<HashTableDTO>();
            try
            {
                var attributeRequest = new RetrieveAttributeRequest
                {
                    EntityLogicalName = "contact",
                    LogicalName = "hil_salutation",
                    RetrieveAsIfPublished = true
                };

                var attributeResponse = (RetrieveAttributeResponse)service.Execute(attributeRequest);
                var attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

                var optionList = (from o in attributeMetadata.OptionSet.Options
                                  select new { Value = o.Value, Text = o.Label.UserLocalizedLabel.Label }).ToList();
                foreach (var option in optionList)
                {
                    lstIoTSalutationEnum.Add(new HashTableDTO() { Value = option.Value, Label = option.Text, StatusCode = "200", StatusDescription = "OK" });
                }
            }
            catch (Exception ex)
            {
                objIoTSalutationEnum = new HashTableDTO { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                lstIoTSalutationEnum.Add(objIoTSalutationEnum);
            }
            return lstIoTSalutationEnum;
        }
    }
    public class IoTValidateSerialNumber
    {
        public string SerialNumber { get; set; }
        public string ProductCategory { get; set; }
        public Guid ProductCategoryGuid { get; set; }
        public string ProductSubCategory { get; set; }
        public Guid ProductSubCategoryGuid { get; set; }
        public string ModelCode { get; set; }
        public string ModelName { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
        public Guid? ProductId { get; set; }
    }
}

