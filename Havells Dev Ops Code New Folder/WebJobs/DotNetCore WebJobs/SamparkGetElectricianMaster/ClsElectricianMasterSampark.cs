using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System.Configuration;
using System.Text;
using System.Text.RegularExpressions;

namespace SamparkGetElectricianMaster
{
    public class ClsElectricianMasterSampark
    {
        private readonly ServiceClient service;
        public ClsElectricianMasterSampark(ServiceClient _service)
        {
            service = _service;
        }
        public void GetElectricianMasterForSampark()
        {
            Console.WriteLine($"******************************* syncSamparkElectricianMaster Started on {DateTime.Now} ******************************* ");

            IntegrationConfiguration integrationConfiguration = GetIntegrationConfiguration("SamparkGetElectricianMaster");
            string authInfo = integrationConfiguration.UserName + ":" + integrationConfiguration.Password;
            authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
            //string enquiryDate = DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "000000"; // Yesterday's date
            //string todate = DateTime.Now.ToString("yyyyMMdd") + "000000"; // Today's date

            //Console.WriteLine($"************************* WebJob Execution Summary (FromDate: {enquiryDate}, ToDate: {todate}) *************************");
            //string queryString = $"enquiryDate={enquiryDate}&todate={todate}";

            string[] Branches = ConfigurationManager.AppSettings["Branches"].Split(";");
            foreach (string Branch in Branches)
            {
                Console.WriteLine($"******************************* Processing with Branch code: {Branch} *******************************");

                string uri = integrationConfiguration.Url;
                string queryString = $"Branch={Branch}";
                uri += queryString;
                Console.WriteLine("Final URL: " + uri);
                var client = new RestClient(uri);
                var request = new RestRequest();
                request.AddHeader("Authorization", "Basic " + authInfo);
                RestResponse response = client.Execute(request, Method.Get);
                ElectricianrootObject responseData = JsonConvert.DeserializeObject<ElectricianrootObject>(response.Content);

                int iDone = 0;
                int iTotal = responseData.Results.Count;
                Console.WriteLine("Total records to process: " + iTotal);
                foreach (ElectricianModel objitem in responseData.Results)
                {
                    //Console.WriteLine("Print Api Response output: " + JsonConvert.SerializeObject(objitem));
                    Console.WriteLine($"Processing record {iDone + 1} out of {iTotal}");
                    try
                    {
                        iDone++;
                        Guid accountId = Guid.Empty;
                        string _accountFetch = null;
                        Entity accountEntity = new Entity("account");
                        EntityCollection accountEntities = null;
                        // Query to check if Electrician MDM Code exists in the "account" table
                        if (!string.IsNullOrWhiteSpace(objitem.MDMUserCode))
                        {
                            _accountFetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='account'>
                                <attribute name='accountid'/>
                                <attribute name='hil_vendorcode'/>
                                <attribute name='accountnumber'/>
                                <attribute name='createdon'/>
                                <attribute name='customertypecode' />
                                <filter type='or'>
                                    <condition attribute='msdyn_externalaccountid' operator='eq' value='{objitem.MDMUserCode}'/>
                                    <condition attribute='hil_vendorcode' operator='eq' value='{objitem.MDMUserCode}'/>
                                    <condition attribute='accountnumber' operator='eq' value='{objitem.MDMUserCode}'/>
                                </filter>
                                </entity>
                                </fetch>";
                            accountEntities = service.RetrieveMultiple(new FetchExpression(_accountFetch));
                            if (accountEntities.Entities.Count > 0)
                            {
                                try
                                {
                                    if (accountEntities.Entities[0].Contains("customertypecode"))
                                    {

                                        OptionSetValue _customerType = accountEntities.Entities[0].GetAttributeValue<OptionSetValue>("customertypecode");
                                        if (_customerType.Value == 6 || _customerType.Value == 9) //Franchise || DSE
                                        {
                                            Console.WriteLine($"{objitem.MDMUserCode} Sampark SAP Code already exists in D365 as Frachise/DSE.");
                                            continue;
                                        }
                                    }

                                    Console.WriteLine($"{objitem.MDMUserCode} Sampark Electrician MDM Code already exists in D365");

                                    accountId = accountEntities.Entities[0].Id;
                                    accountEntity.Id = accountId;
                                    accountEntity["hil_vendorcode"] = null;
                                    accountEntity["accountnumber"] = objitem.MDMUserCode.Trim();
                                    accountEntity["msdyn_externalaccountid"] = objitem.MDMUserCode.Trim();
                                    accountEntity["address1_telephone2"] = IsValidMobileNumber(objitem.ElectricianCode);
                                    GetChannelPartnerType(service, objitem.UserType, objitem.UserSubType, accountEntity);
                                    UpdateChannelPartner(service, objitem, accountEntity);
                                    UpdateAccountStatus(accountEntity, objitem.IsActive);
                                    try
                                    {
                                        if (objitem.UserType.ToUpper() == "DEALER")
                                            accountEntity["accountnumber"] = objitem.ElectricianSamparkCode; //Sap Code

                                        service.Update(accountEntity);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"ERROR! {objitem.MDMUserCode} {ex.Message} {ex.StackTrace}");
                                    }

                                    Console.WriteLine("Sampark Electrician MDM Code updated sucessfully in D365 " + objitem.MDMUserCode);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"ERROR!!! Processing record {iDone}: {ex.Message.ToUpper()}");
                                }
                            }
                            else
                            {
                                CreateChannelPartner(service, objitem, accountEntity);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"ERROR!!! Processing record {iDone}: {ex.Message}{ex.StackTrace}");
                    }
                }
                Console.WriteLine($"All records processed. Total records: {iTotal}, Successfully processed: {iDone}");

                Console.WriteLine("******************************* syncSamparkElectricianMaster Ended *******************************");
            }
        }
        private static void CreateChannelPartner(IOrganizationService service, ElectricianModel objitem, Entity accountEntity)
        {
            string businessMappingFetch = string.Empty;
            accountEntity["name"] = !string.IsNullOrWhiteSpace(objitem.ElectricianFirmName) ? objitem.ElectricianFirmName : string.Empty;
            accountEntity["hil_retailerfullname"] = !string.IsNullOrWhiteSpace(objitem.ElectricianFullName) ? objitem.ElectricianFullName : string.Empty;
            accountEntity["emailaddress1"] = !string.IsNullOrWhiteSpace(objitem.ElectricianEmail) ? objitem.ElectricianEmail : string.Empty;
            accountEntity["telephone1"] = IsValidMobileNumber(objitem.ElectricianCode);
            accountEntity["address1_telephone2"] = IsValidMobileNumber(objitem.ElectricianCode);
            accountEntity["address1_line1"] = !string.IsNullOrWhiteSpace(objitem.Address1) ? objitem.Address1 : string.Empty;
            accountEntity["address1_line2"] = !string.IsNullOrWhiteSpace(objitem.Address2) ? objitem.Address2 : string.Empty;
            accountEntity["hil_pan"] = !string.IsNullOrWhiteSpace(objitem.PAN) ? objitem.PAN : string.Empty;
            accountEntity["address1_postofficebox"] = !string.IsNullOrWhiteSpace(objitem.GSTNumber) ? objitem.GSTNumber : string.Empty;

            GetChannelPartnerType(service, objitem.UserType, objitem.UserSubType, accountEntity);

            businessMappingFetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_businessmapping'>
                <attribute name='hil_businessmappingid' />
                <attribute name='hil_name' />
                <attribute name='hil_stagingarea' />
                <attribute name='hil_stagingpin' />
                <attribute name='hil_subterritory' />
                <attribute name='hil_state' />
                <attribute name='hil_salesoffice' />
                <attribute name='hil_region' />
                <attribute name='hil_pincode' />
                <attribute name='hil_district' />
                <attribute name='hil_city' />
                <attribute name='hil_branch' />
                <attribute name='hil_area' />
                <filter type='and'>
                <condition attribute='hil_stagingarea' operator='eq' value='{objitem.AreaCode}' />
                <condition attribute='hil_stagingpin' operator='eq' value='{objitem.PINCode}' />
                <condition attribute='statecode' operator='eq' value='0' />
                </filter>
                </entity>
                </fetch>";
            EntityCollection businessMappingEntities = service.RetrieveMultiple(new FetchExpression(businessMappingFetch));
            if (businessMappingEntities.Entities.Count > 0)
            {
                Entity businessMapping = businessMappingEntities.Entities[0];
                if (businessMapping.Contains("hil_pincode"))
                    accountEntity["hil_pincode"] = businessMapping["hil_pincode"];
                if (businessMapping.Contains("hil_state"))
                    accountEntity["hil_state"] = businessMapping["hil_state"];
                if (businessMapping.Contains("hil_district"))
                    accountEntity["hil_district"] = businessMapping["hil_district"];
                if (businessMapping.Contains("hil_city"))
                    accountEntity["hil_city"] = businessMapping["hil_city"];
                if (businessMapping.Contains("hil_branch"))
                    accountEntity["hil_branch"] = businessMapping["hil_branch"];
                if (businessMapping.Contains("hil_salesoffice"))
                    accountEntity["hil_salesoffice"] = businessMapping["hil_salesoffice"];
                if (businessMapping.Contains("hil_region"))
                    accountEntity["hil_region"] = businessMapping["hil_region"];
                if (businessMapping.Contains("hil_area"))
                    accountEntity["hil_area"] = businessMapping["hil_area"];
                if (businessMapping.Contains("hil_subterritory"))
                    accountEntity["hil_subterritory"] = businessMapping["hil_subterritory"];
            }
            else
            {
                Console.WriteLine($"BusinessMapping for AreaCode {objitem.AreaCode} and PINCode {objitem.PINCode} does not exist in D365.");
            }
            UpdateAccountStatus(accountEntity, objitem.IsActive);
            accountEntity["hil_vendorcode"] = null;
            accountEntity["accountnumber"] = objitem.MDMUserCode.Trim();
            accountEntity["msdyn_externalaccountid"] = objitem.MDMUserCode.Trim();

            if (objitem.UserType.ToUpper() == "DEALER")
            {
                accountEntity["accountnumber"] = objitem.ElectricianSamparkCode; // Sap Code
            }
            try
            {
                Guid guid = service.Create(accountEntity);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR! {objitem.MDMUserCode} {ex.Message} {ex.StackTrace}");
            }
            Console.WriteLine($"{objitem.MDMUserCode} Sampark ElectricianMaster created successfully in D365.");
        }
        private static void UpdateChannelPartner(IOrganizationService service, ElectricianModel objitem, Entity accountEntity)
        {
            accountEntity["name"] = !string.IsNullOrWhiteSpace(objitem.ElectricianFirmName) ? objitem.ElectricianFirmName : string.Empty;
            accountEntity["hil_retailerfullname"] = !string.IsNullOrWhiteSpace(objitem.ElectricianFullName) ? objitem.ElectricianFullName : string.Empty;
            accountEntity["emailaddress1"] = !string.IsNullOrWhiteSpace(objitem.ElectricianEmail) ? objitem.ElectricianEmail : string.Empty;
            accountEntity["address1_telephone2"] = IsValidMobileNumber(objitem.ElectricianCode);
            accountEntity["address1_line1"] = !string.IsNullOrWhiteSpace(objitem.Address1) ? objitem.Address1 : string.Empty;
            accountEntity["address1_line2"] = !string.IsNullOrWhiteSpace(objitem.Address2) ? objitem.Address2 : string.Empty;
            accountEntity["hil_pan"] = !string.IsNullOrWhiteSpace(objitem.PAN) ? objitem.PAN : string.Empty;
            accountEntity["address1_postofficebox"] = !string.IsNullOrWhiteSpace(objitem.GSTNumber) ? objitem.GSTNumber : string.Empty;
        }
        private static void GetChannelPartnerType(IOrganizationService service, string _userType, string _userSubType, Entity _entRecord)
        {
            EntityReference _cpType = null;
            EntityReference _cpSubType = null;
            if (!string.IsNullOrEmpty(_userType))
            {
                var fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_usertype'>
                <attribute name='hil_usertypeid' />
                <filter type='and'>
                    <condition attribute='statecode' operator='eq' value='0' />
                    <condition attribute='hil_name' operator='eq' value='{_userType.ToUpper()}' />
                    <condition attribute='hil_parentusertype' operator='null' />
                </filter>
                </entity>
                </fetch>";
                var entColl = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (entColl.Entities.Count > 0)
                {
                    _cpType = entColl.Entities[0].ToEntityReference();
                }
            }
            if (!string.IsNullOrEmpty(_userSubType) && _cpType != null)
            {
                var fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_usertype'>
                <attribute name='hil_usertypeid' />
                <filter type='and'>
                    <condition attribute='statecode' operator='eq' value='0' />
                    <condition attribute='hil_name' operator='eq' value='{_userSubType.ToUpper()}' />
                    <condition attribute='hil_parentusertype' operator='eq' value='{_cpType.Id}' />
                </filter>
                </entity>
                </fetch>";
                var entColl = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (entColl.Entities.Count > 0)
                {
                    _cpSubType = entColl.Entities[0].ToEntityReference();
                }
            }
            _entRecord["hil_usertype"] = _cpType;
            _entRecord["hil_usersubtype"] = _cpSubType;
        }
        private static void UpdateAccountStatus(Entity accountEntity, bool isActive)
        {
            if (isActive)
            {
                accountEntity["hil_samparkmasterstatus"] = new OptionSetValue(0); // Active
            }
            else
            {
                accountEntity["hil_samparkmasterstatus"] = new OptionSetValue(1); // Inactive
            }
        }
        private IntegrationConfiguration GetIntegrationConfiguration(string APIName)
        {
            try
            {
                IntegrationConfiguration inconfig = new IntegrationConfiguration();
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, APIName);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                inconfig.Url = integrationConfiguration.GetAttributeValue<string>("hil_url");
                inconfig.UserName = integrationConfiguration.GetAttributeValue<string>("hil_username");
                inconfig.Password = integrationConfiguration.GetAttributeValue<string>("hil_password");
                return inconfig;
            }
            catch (Exception ex)
            {
                throw new Exception("Error : " + ex.Message);
            }
        }
        private static string? IsValidMobileNumber(string mobileNumber)
        {
            // Regex pattern to match numbers starting with 6 and having exactly 10 digits
            string pattern = @"^[6-9]\d{9}$";
            if (!string.IsNullOrEmpty(mobileNumber) && Regex.IsMatch(mobileNumber, pattern))
            {
                return mobileNumber;
            }
            return null;
        }
    }
    public class ElectricianrootObject
    {
        public string Result { get; set; }
        public List<ElectricianModel> Results { get; set; }
    }
    public class ElectricianModel
    {
        public string ElectricianCode { get; set; }
        public string MDMUserCode { get; set; }
        public string ElectricianSamparkCode { get; set; }
        public string ElectricianFullName { get; set; }
        public string ElectricianFirmName { get; set; }
        public string ElectricianEmail { get; set; }
        public string PAN { get; set; }
        public string GSTNumber { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string AreaCode { get; set; }
        public string PINCode { get; set; }
        public bool IsActive { get; set; }
        public string UserType { get; set; }
        public string UserSubType { get; set; }
    }
    public class IntegrationConfiguration
    {
        public string Url { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
    }
}
