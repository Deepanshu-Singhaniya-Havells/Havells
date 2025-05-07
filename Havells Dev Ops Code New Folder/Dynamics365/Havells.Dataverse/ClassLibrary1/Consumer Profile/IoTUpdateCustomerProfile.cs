using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Havells.Dataverse.CustomConnector.Consumer_Profile
{

    public class IoTUpdateCustomerProfile : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            string SalutationNameRegex = @"^[a-zA-Z./]+$";
            try
            {
                IoTCustomerProfile customerProfile = new IoTCustomerProfile();

                #region extract
                string CustomerGuid = Convert.ToString(context.InputParameters["CustomerGuid"]);
                string MobileNumber = Convert.ToString(context.InputParameters["MobileNumber"]);
                string SalutationName = Convert.ToString(context.InputParameters["SalutationName"]);
                int Salutation = Convert.ToInt32(context.InputParameters["Salutation"]);
                string FirstName = Convert.ToString(context.InputParameters["FirstName"]);
                string MiddleName = Convert.ToString(context.InputParameters["MiddleName"]);
                string LastName = Convert.ToString(context.InputParameters["LastName"]);
                bool Gender = Convert.ToBoolean(context.InputParameters["Gender"]);
                string GenderName = Convert.ToString(context.InputParameters["GenderName"]);
                string DateOfBirth = Convert.ToString(context.InputParameters["DateOfBirth"]);
                string DateOfAnniversary = Convert.ToString(context.InputParameters["DateOfAnniversary"]);
                string Email = Convert.ToString(context.InputParameters["Email"]);
                string AlternateNumber = Convert.ToString(context.InputParameters["AlternateNumber"]);
                bool Consent = Convert.ToBoolean(context.InputParameters["Consent"]);
                bool SubscribeForMsgService = Convert.ToBoolean(context.InputParameters["SubscribeForMsgService"]);
                string PreferredLanguage = Convert.ToString(context.InputParameters["PreferredLanguage"]);
                int CustomerType = Convert.ToInt32(context.InputParameters["CustomerType"]);
                bool IsPremiumCustomer = Convert.ToBoolean(context.InputParameters["IsPremiumCustomer"]);

                #endregion

                Guid _CustomerGuid;

                if (string.IsNullOrWhiteSpace(CustomerGuid))
                {
                    customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "CustomerGUID is required." };
                    string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                    context.OutputParameters["data"] = jsonResultResponse;
                    return;
                }
                if (Guid.TryParse(CustomerGuid, out _CustomerGuid))
                {
                    if (APValidate.IsvalidGuid(CustomerGuid))
                    {
                        customerProfile.CustomerGuid = _CustomerGuid; // (1)
                    }
                }
                else
                {
                    customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "CustomerGUID is not valid." };
                    string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                    context.OutputParameters["data"] = jsonResultResponse;
                    return;
                }
                if (!string.IsNullOrWhiteSpace(MobileNumber))
                {
                    if (!APValidate.IsValidMobileNumber(MobileNumber))
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "MobileNumber is not valid." };
                        string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                        context.OutputParameters["data"] = jsonResultResponse;
                        return;
                    }
                    customerProfile.MobileNumber = MobileNumber; // (2)
                }
                if (!string.IsNullOrWhiteSpace(PreferredLanguage))
                {
                    if (!APValidate.IsValidString(PreferredLanguage))
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "PreferredLanguage is not valid, should be valid alphabet string." };
                        string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                        context.OutputParameters["data"] = jsonResultResponse;
                        return;
                    }
                    customerProfile.PreferredLanguage = PreferredLanguage; // (3)
                }

                if (!string.IsNullOrWhiteSpace(SalutationName))
                {
                    // Validate SalutationName based on Salutation value
                    string[] validSalutationNames = { "Miss", "Mr.", "Mrs.", "Dr.", "M/S" };
                    int salutationValue = int.Parse(Salutation.ToString());
                    if (salutationValue >= 1 && salutationValue <= 5)
                    {
                        string expectedSalutationName = validSalutationNames[salutationValue - 1];
                        if (SalutationName != expectedSalutationName)
                        {
                            customerProfile = new IoTCustomerProfile
                            {
                                StatusCode = "204",
                                StatusDescription = $"For Salutation {salutationValue}, the SalutationName must be - " + expectedSalutationName
                            };
                            string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                            context.OutputParameters["data"] = jsonResultResponse;
                            return;
                        }
                    }
                    if (!Regex.IsMatch(SalutationName, SalutationNameRegex))
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "SalutationName is not Valid. only Alphabets and {.,/} are allowed" };
                        string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                        context.OutputParameters["data"] = jsonResultResponse;
                        return;
                    }
                    customerProfile.SalutationName = SalutationName; // (5)
                }

                if (!string.IsNullOrWhiteSpace(FirstName))
                {
                    string pattern = @"^[a-zA-Z.]*$";
                    if (!Regex.IsMatch(FirstName, pattern))
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "FirstName is not Valid" };
                        string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                        context.OutputParameters["data"] = jsonResultResponse;
                        return;
                    }
                    customerProfile.FirstName = FirstName; // (6)
                }
                if (!string.IsNullOrWhiteSpace(MiddleName))
                {
                    string pattern = @"^[a-zA-Z.]*$";
                    if (!Regex.IsMatch(MiddleName, pattern))
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "MiddleName is not Valid" };
                        string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                        context.OutputParameters["data"] = jsonResultResponse;
                        return;
                    }
                    customerProfile.MiddleName = MiddleName; // (7)
                }
                if (!string.IsNullOrWhiteSpace(LastName))
                {
                    string pattern = @"^[a-zA-Z.]*$";
                    if (!Regex.IsMatch(LastName, pattern))
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "LastName is not Valid" };
                        string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                        context.OutputParameters["data"] = jsonResultResponse;
                        return;
                    }
                    customerProfile.LastName = LastName; // (8)
                }
               
                if (APValidate.IsValidboolen(Gender.ToString()))
                {
                    customerProfile.Gender = Gender; // (9)
                }
                else
                {
                    customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Gender is not a Valid Boolean {true/false}" };
                    string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                    context.OutputParameters["data"] = jsonResultResponse;
                    return;
                }

                if (!string.IsNullOrWhiteSpace(GenderName))
                {
                    if (Gender)
                    {
                        if (GenderName != "Male")
                        {
                            customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "GenderName is invalid. For Gender value 'true', GenderName must be 'Male'." };
                            string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                            context.OutputParameters["data"] = jsonResultResponse;
                            return;
                        }
                    }
                    else
                    {
                        if (GenderName != "Female")
                        {
                            customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "GenderName is invalid. For Gender value 'false', GenderName must be 'Female'." };
                            string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                            context.OutputParameters["data"] = jsonResultResponse;
                            return;
                        }
                    }
                    customerProfile.GenderName = GenderName; // (10)
                }
                if (!string.IsNullOrWhiteSpace(DateOfBirth))
                {
                    string pattern = @"^(0[1-9]|1[0-2])([./-])([0-2][0-9]|3[01])\2\d{4}$";
                    bool isMatch = Regex.IsMatch(DateOfBirth, pattern);
                    if (isMatch)
                    {
                        customerProfile.DateOfBirth = DateOfBirth; // (11)
                    }
                    else
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "DateOfBirth is not valid.please try with {MM/dd/yyyy} or {MM-dd-yyyy} or {MM.dd.yyyy} format" };
                        string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                        context.OutputParameters["data"] = jsonResultResponse;
                        return;
                    }
                    // Additional validation to ensure the date is logical (e.g., 31-02-2024 is invalid)
                    try
                    {
                        string[] formats = { "MM.dd.yyyy", "MM-dd-yyyy", "MM/dd/yyyy" };
                        DateTime.ParseExact(DateOfBirth, formats, CultureInfo.InvariantCulture, DateTimeStyles.None);
                    }
                    catch (FormatException)
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Invalid DateOfBirth format. It should be valid according to the Gregorian calendar" };
                        string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                        context.OutputParameters["data"] = jsonResultResponse;
                        return;
                    }
                }
                if (!string.IsNullOrWhiteSpace(DateOfAnniversary))
                {
                    string pattern = @"^(0[1-9]|1[0-2])([./-])([0-2][0-9]|3[01])\2\d{4}$";
                    bool isMatch = Regex.IsMatch(DateOfAnniversary, pattern);
                    if (isMatch)
                    {
                        customerProfile.DateOfAnniversary = DateOfAnniversary; // (12)
                    }
                    else
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "DateOfAnniversary is not valid.please try with {MM/dd/yyyy} or {MM-dd-yyyy} or {MM.dd.yyyy} format" };
                        string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                        context.OutputParameters["data"] = jsonResultResponse;
                        return;
                    }
                    // Additional validation to ensure the date is logical (e.g., 30-02-2024 is invalid)
                    try
                    {
                        string[] formats = { "MM.dd.yyyy", "MM-dd-yyyy", "MM/dd/yyyy" };
                        DateTime.ParseExact(DateOfBirth, formats, CultureInfo.InvariantCulture, DateTimeStyles.None);
                    }
                    catch (FormatException)
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Invalid DateOfAnniversary format. It should be valid according to the Gregorian calendar" };
                        string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                        context.OutputParameters["data"] = jsonResultResponse;
                        return;
                    }
                }
                if (!string.IsNullOrWhiteSpace(Email))
                {
                    if (!APValidate.IsValidEmail(Email))
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Email is Not Valid." };
                        string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                        context.OutputParameters["data"] = jsonResultResponse;
                        return;
                    }
                    customerProfile.Email = Email; // (13)
                }
                if (!string.IsNullOrWhiteSpace(AlternateNumber))
                {
                    if (!APValidate.IsValidMobileNumber(AlternateNumber))
                    {
                        customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "AlternateNumber is Not Valid." };
                        string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                        context.OutputParameters["data"] = jsonResultResponse;
                        return;
                    }
                    customerProfile.AlternateNumber = AlternateNumber; // (14)
                }

                if (APValidate.IsValidboolen(Consent.ToString()))
                {
                    customerProfile.Consent = Consent; // (16)
                }
                else
                {
                    customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Consent is not a Valid Boolean {true/false}" };
                    string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                    context.OutputParameters["data"] = jsonResultResponse;
                    return;
                }
                
                if (APValidate.IsValidboolen(SubscribeForMsgService.ToString()))
                {
                    customerProfile.SubscribeForMsgService = SubscribeForMsgService; // (17)
                }
                else
                {
                    customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "SubscribeForMsgService is not a Valid Boolean {true/false}" };
                    string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                    context.OutputParameters["data"] = jsonResultResponse;
                    return;
                }
               
                if (APValidate.IsValidboolen(IsPremiumCustomer.ToString()))
                {
                    customerProfile.IsPremiumCustomer = IsPremiumCustomer; // (18)
                }
                else
                {
                    customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "IsPremiumCustomer is not a Valid Boolean {true/false}" };
                    string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                    context.OutputParameters["data"] = jsonResultResponse;
                    return;
                }

                IoTCustomerProfile objIoTCustomerProfile = new IoTCustomerProfile();

                IoTCustomerProfile ioTCustomerProfile1 = objIoTCustomerProfile.UpdateIoTCustomerProfile(customerProfile, service);
                var serializedResponse = JsonSerializer.Serialize(ioTCustomerProfile1);
                context.OutputParameters["data"] = serializedResponse;
                return;

            }
            catch (Exception ex)
            {
                IoTCustomerProfile customerProfile = new IoTCustomerProfile();
                customerProfile = new IoTCustomerProfile { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message };
                string jsonResultResponse = JsonSerializer.Serialize(customerProfile);
                context.OutputParameters["data"] = jsonResultResponse;
                return;
            }
        }
        public class IoTCustomerProfile
        {
            public Guid CustomerGuid { get; set; }
            public string MobileNumber { get; set; }
            public string SalutationName { get; set; }
            public int? Salutation { get; set; }
            public string FirstName { get; set; }
            public string MiddleName { get; set; }
            public string LastName { get; set; }
            public bool? Gender { get; set; }
            public string GenderName { get; set; }
            public string DateOfBirth { get; set; }
            public string DateOfAnniversary { get; set; }
            public string Email { get; set; }
            public string AlternateNumber { get; set; }
            public bool? Consent { get; set; }
            public bool? SubscribeForMsgService { get; set; }
            public string PreferredLanguage { get; set; }
            public int? CustomerType { get; set; }
            public bool? IsPremiumCustomer { get; set; }
            public string StatusCode { get; set; }
            public string StatusDescription { get; set; }

            public IoTCustomerProfile UpdateIoTCustomerProfile(IoTCustomerProfile customerProfileData, IOrganizationService service)
            {
                IoTCustomerProfile customerProfile;
                try
                {
                    EntityReference _entPreferredLang = null;

                    if (service != null)
                    {
                        if (customerProfileData.CustomerGuid.ToString().Trim().Length == 0)
                        {
                            customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Customer GUID is required." };
                            return customerProfile;
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
                                customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Preferred Language does not exist." };
                                return customerProfile;
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
                            Entity entCustomer = new Entity("contact", ent.Id);
                            if (customerProfileData.Salutation != null)
                            {
                                List<HashTableDTO> salutationList = new List<HashTableDTO>();
                                IoTCommonLibUpdateProfile commonLib = new IoTCommonLibUpdateProfile();
                                if (!commonLib.GetSalutationEnum(service).Exists(x => x.Value == customerProfileData.Salutation))
                                {
                                    customerProfile = new IoTCustomerProfile { StatusCode = "204", StatusDescription = "Salutation not found in D365. To get list of valid Salutations,Please Call {GetSalutationEnum} API and pass proper Salutation." };
                                    return customerProfile;
                                }
                                entCustomer["hil_salutation"] = new OptionSetValue(customerProfileData.Salutation.Value);
                            }
                            if (customerProfileData.FirstName != null)
                                entCustomer["firstname"] = customerProfileData.FirstName;
                            if (customerProfileData.MiddleName != null)
                                entCustomer["middlename"] = customerProfileData.MiddleName;
                            if (customerProfileData.LastName != null)
                                entCustomer["lastname"] = customerProfileData.LastName;
                            if (customerProfileData.AlternateNumber != null)
                                entCustomer["address1_telephone3"] = customerProfileData.AlternateNumber;
                            if (customerProfileData.Gender != null)
                                entCustomer["hil_gender"] = !customerProfileData.Gender;

                            if (customerProfileData.Email != null)
                                entCustomer["emailaddress1"] = customerProfileData.Email;

                            if (customerProfileData.DateOfBirth != null)
                            {
                                DateTime dtDOB = DateTime.ParseExact(customerProfileData.DateOfBirth, new string[] { "MM.dd.yyyy", "MM-dd-yyyy", "MM/dd/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None);
                                entCustomer["hil_dob"] = dtDOB;
                            }
                            if (customerProfileData.DateOfAnniversary != null)
                            {
                                DateTime dtDOA = DateTime.ParseExact(customerProfileData.DateOfAnniversary, new string[] { "MM.dd.yyyy", "MM-dd-yyyy", "MM/dd/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None);
                                entCustomer["hil_doa"] = dtDOA;
                            }
                            if (customerProfileData.Consent != null)
                                entCustomer["hil_consent"] = customerProfileData.Consent;

                            if (customerProfileData.SubscribeForMsgService != null)
                                entCustomer["hil_subscribeformessagingservice"] = customerProfileData.SubscribeForMsgService;

                            if (_entPreferredLang != null)
                                entCustomer["hil_preferredlanguageforcommunication"] = _entPreferredLang;

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
            public class HashTableDTO
            {
                public string Label { get; set; }
                public int? Value { get; set; }
                public string Extension { get; set; }
                public string StatusCode { get; set; }
                public string StatusDescription { get; set; }
            }
        }
        public class IoTCommonLibUpdateProfile
        {
            public List<HashTableDTO> GetSalutationEnum(IOrganizationService service)
            {
                HashTableDTO objIoTSalutationEnum;
                List<HashTableDTO> lstIoTSalutationEnum = new List<HashTableDTO>();
                try
                {
                    if (service != null)
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
                    else
                    {
                        objIoTSalutationEnum = new HashTableDTO { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                        lstIoTSalutationEnum.Add(objIoTSalutationEnum);
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
    }
}

