using Microsoft.Crm.Sdk.Messages;
using Microsoft.Identity.Client;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Configuration;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace ImportSamparkChannelPartner
{
    public class ClsRetailerMasterSampark
    {
        private readonly ServiceClient service;
        public ClsRetailerMasterSampark(ServiceClient _service)
        {
            service = _service;
        }
        public void GetRetailerMasterSampark()
        {
            Console.WriteLine($"******************************* syncSamparkRetailerMaster Started on {DateTime.Now}  ******************************* ");

            GetChannelPartnerModel objparamreq = new GetChannelPartnerModel();
            IntegrationConfiguration integrationConfiguration = GetIntegrationConfiguration(service, "SamparkGetRetailersMaster");
            //DateTime FromDate = DateTime.Now.AddDays(-1);
            //objparamreq.FromDate = FromDate.Year.ToString() + FromDate.Month.ToString().PadLeft(2, '0') + FromDate.Day.ToString().PadLeft(2, '0') + "000000";
            //objparamreq.ToDate = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0') + "000000";
            //Console.WriteLine($"************************* WebJob Execution Summary (FromDate: {objparamreq.FromDate}, ToDate: {objparamreq.ToDate}) *************************");

            string[] Branches = ConfigurationManager.AppSettings["Branches"].Split(";");

            foreach (string Branch in Branches)
            {
                objparamreq.Branch = Branch;

                Console.WriteLine($"******************************* Processing with Branch code: {Branch} *******************************");

                ChannelPartnerList response = JsonConvert.DeserializeObject<ChannelPartnerList>(CallAPI(integrationConfiguration, JsonConvert.SerializeObject(objparamreq), "POST"));
                int iDone = 0;
                int iTotal = response.Results.Count;
                Console.WriteLine("Total records to process: " + iTotal);
                foreach (ChannelPartnerModel objitem in response.Results)
                {
                    Console.WriteLine($"Processing record {iDone + 1} out of {iTotal}");
                    try
                    {
                        iDone++;
                        Guid accountId = Guid.Empty;
                        string _accountFetch = null;
                        Entity accountEntity = new Entity("account");
                        EntityCollection accountEntities = null;
                        try
                        {
                            //Query to check if RetailerSAPCode exists in the "account" table
                            if (!string.IsNullOrWhiteSpace(objitem.RetailerSAPCode))
                            {
                                _accountFetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='account'>
                                    <attribute name='accountid' />
                                    <attribute name='customertypecode' />
                                    <filter type='or'>
                                    <condition attribute='accountnumber' operator='eq' value='{objitem.RetailerSAPCode.Trim()}' />
                                    </filter>
                                    </entity>
                                    </fetch>";
                                accountEntities = service.RetrieveMultiple(new FetchExpression(_accountFetch));
                                if (accountEntities.Entities.Count > 0)
                                {
                                    if (accountEntities.Entities[0].Contains("customertypecode"))
                                    {
                                        OptionSetValue _customerType = accountEntities.Entities[0].GetAttributeValue<OptionSetValue>("customertypecode");
                                        if (_customerType.Value == 6 || _customerType.Value == 9) //Franchise || DSE
                                        {
                                            Console.WriteLine($"{objitem.RetailerSAPCode} Sampark SAP Code already exists in D365 as Frachise/DSE.");
                                            continue;
                                        }
                                    }
                                    Console.WriteLine($"{objitem.RetailerSAPCode} Sampark SAP Code already exists in D365.");
                                    accountId = accountEntities.Entities[0].Id;
                                    accountEntity.Id = accountId;
                                    accountEntity["hil_vendorcode"] = null;
                                    accountEntity["msdyn_externalaccountid"] = objitem.RetailerCode.Trim();
                                    UpdateChannelPartner(service, objitem, accountEntity);
                                    GetChannelPartnerType(service, objitem.UserType, objitem.UserSubType, accountEntity);
                                    UpdateAccountStatus(accountEntity, objitem.IsActive);
                                    try
                                    {
                                        service.Update(accountEntity);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"ERROR! {objitem.RetailerSAPCode} {ex.Message} {ex.StackTrace}");
                                    }

                                    Console.WriteLine($"Update done SAP code: {objitem.RetailerSAPCode} with MDM code {objitem.RetailerCode}");
                                }
                                else
                                {
                                    // Check fallback conditions if RetailerSAPCode does not exist
                                    _accountFetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='account'>
                                        <attribute name='accountid' />
                                        <attribute name='hil_vendorcode' />
                                        <attribute name='accountnumber' />
                                        <attribute name='customertypecode' />
                                        <attribute name='createdon' />
                                        <filter type='or'>
                                        <condition attribute='msdyn_externalaccountid' operator='eq' value='{objitem.RetailerCode.Trim()}' />
                                        <condition attribute='hil_vendorcode' operator='eq' value='{objitem.RetailerCode.Trim()}' />
                                        </filter>
                                        </entity>
                                        </fetch>";
                                    accountEntities = service.RetrieveMultiple(new FetchExpression(_accountFetch));
                                    if (accountEntities.Entities.Count > 0)
                                    {
                                        try
                                        {
                                            Console.WriteLine($"Sampark Retailer/Dealer MDM Code already exists in D365.");

                                            accountId = accountEntities.Entities[0].Id;
                                            accountEntity.Id = accountId;
                                            accountEntity["hil_vendorcode"] = null;
                                            accountEntity["msdyn_externalaccountid"] = objitem.RetailerCode.Trim();
                                            UpdateChannelPartner(service, objitem, accountEntity);
                                            GetChannelPartnerType(service, objitem.UserType, objitem.UserSubType, accountEntity);
                                            UpdateAccountStatus(accountEntity, objitem.IsActive);
                                            try
                                            {
                                                if (accountEntities.Entities[0].Contains("customertypecode"))
                                                {
                                                    OptionSetValue _customerType = accountEntities.Entities[0].GetAttributeValue<OptionSetValue>("customertypecode");
                                                    if (_customerType.Value == 6 || _customerType.Value == 9) //Franchise || DSE
                                                    {
                                                        Console.WriteLine($"{objitem.RetailerSAPCode} Sampark SAP Code already exists in D365 as Frachise/DSE.");
                                                        continue;
                                                    }
                                                }
                                                if (objitem.UserType.ToUpper() == "DEALER")
                                                {
                                                    accountEntity["accountnumber"] = objitem.RetailerSAPCode;
                                                }
                                                else
                                                {
                                                    if (!accountEntities.Entities[0].Contains("accountnumber"))
                                                    {
                                                        accountEntity["accountnumber"] = objitem.RetailerCode.Trim();
                                                    }

                                                }
                                                service.Update(accountEntity);
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"ERROR! {objitem.RetailerCode} {ex.Message} {ex.StackTrace}");
                                            }

                                            Console.WriteLine("Sampark Retailer Master updated successfully in D365.");
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"ERROR!!! Processing record {iDone}: {ex.Message.ToUpper()}");
                                        }
                                    }
                                    else
                                    {
                                        // Create a new channel partner if no records exist
                                        Console.WriteLine($"No existing records found for {objitem.RetailerCode}. Creating a new account.");
                                        CreateChannelPartner(service, objitem, accountEntity);
                                    }
                                }
                            }
                            else
                            {
                                // Direct fallback if RetailerSAPCode is not present
                                _accountFetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='account'>
                                    <attribute name='accountid' />
                                    <attribute name='hil_vendorcode' />
                                    <attribute name='accountnumber' />
                                    <attribute name='createdon' />
                                    <attribute name='customertypecode' />
                                    <filter type='or'>
                                    <condition attribute='msdyn_externalaccountid' operator='eq' value='{objitem.RetailerCode.Trim()}' />
                                    <condition attribute='hil_vendorcode' operator='eq' value='{objitem.RetailerCode.Trim()}' />
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
                                                Console.WriteLine($"{objitem.RetailerCode} Sampark SAP Code already exists in D365 as Frachise/DSE.");
                                                continue;
                                            }
                                        }
                                        Console.WriteLine($"{objitem.RetailerCode} Sampark Retailer/Dealer Code already exists in D365.");

                                        accountId = accountEntities.Entities[0].Id;
                                        accountEntity.Id = accountId;
                                        accountEntity["hil_vendorcode"] = null;
                                        accountEntity["msdyn_externalaccountid"] = objitem.RetailerCode.Trim();
                                        UpdateChannelPartner(service, objitem, accountEntity);
                                        GetChannelPartnerType(service, objitem.UserType, objitem.UserSubType, accountEntity);
                                        UpdateAccountStatus(accountEntity, objitem.IsActive);
                                        try
                                        {
                                            service.Update(accountEntity);
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"ERROR! {objitem.RetailerCode} {ex.Message} {ex.StackTrace}");
                                        }

                                        Console.WriteLine($"{objitem.RetailerCode} Sampark Retailer Master updated successfully in D365.");
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"ERROR!!! Processing record {iDone}: {ex.Message.ToUpper()}");
                                    }
                                }
                                else
                                {
                                    // Create new channel Partner when no records exist for fallback
                                    Console.WriteLine($"No existing records found for {objitem.RetailerCode}. Creating a new account.");
                                    CreateChannelPartner(service, objitem, accountEntity);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"ERROR!!! Processing record {iDone}: {ex.Message}{ex.StackTrace}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"ERROR!!! Processing record {iDone}: {ex.Message}{ex.StackTrace}");
                    }
                }
                Console.WriteLine($"All records processed. Total records: {iTotal}, Successfully processed: {iDone}");

                Console.WriteLine("******************************* syncSamparkRetailerMaster Ended *******************************");
            }
        }
        private static void CreateChannelPartner(IOrganizationService service, ChannelPartnerModel objitem, Entity accountEntity)
        {
            string businessMappingFetch;
            accountEntity["name"] = !string.IsNullOrWhiteSpace(objitem.RetailerFirmName) ? objitem.RetailerFirmName : string.Empty;
            accountEntity["hil_retailerfullname"] = !string.IsNullOrWhiteSpace(objitem.RetailerFullName) ? objitem.RetailerFullName : string.Empty;
            accountEntity["emailaddress1"] = !string.IsNullOrWhiteSpace(objitem.RetailerEmail) ? objitem.RetailerEmail : string.Empty;
            accountEntity["telephone1"] = IsValidMobileNumber(objitem.s_MobileNo);
            accountEntity["address1_telephone2"] = IsValidMobileNumber(objitem.s_MobileNo);
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
            accountEntity["msdyn_externalaccountid"] = objitem.RetailerCode.Trim();

            if (objitem.UserType.ToUpper() == "DEALER")
                accountEntity["accountnumber"] = objitem.RetailerSAPCode;
            else
                accountEntity["accountnumber"] = objitem.RetailerCode.Trim();
            try
            {
                Guid guid = service.Create(accountEntity);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR! {objitem.RetailerCode} {ex.Message} {ex.StackTrace}");
            }

            Console.WriteLine($"{objitem.RetailerCode} Sampark Retailer Master created successfully in D365.");
        }
        private static void UpdateChannelPartner(IOrganizationService service, ChannelPartnerModel objitem, Entity accountEntity)
        {
            accountEntity["hil_retailerfullname"] = !string.IsNullOrWhiteSpace(objitem.RetailerFullName) ? objitem.RetailerFullName : string.Empty;
            accountEntity["emailaddress1"] = !string.IsNullOrWhiteSpace(objitem.RetailerEmail) ? objitem.RetailerEmail : string.Empty;
            accountEntity["address1_telephone2"] = IsValidMobileNumber(objitem.s_MobileNo);
            accountEntity["address1_line1"] = !string.IsNullOrWhiteSpace(objitem.Address1) ? objitem.Address1 : string.Empty;
            accountEntity["address1_line2"] = !string.IsNullOrWhiteSpace(objitem.Address2) ? objitem.Address2 : string.Empty;
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
        private static IntegrationConfiguration GetIntegrationConfiguration(IOrganizationService service, string APIName)
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
                inconfig.url = integrationConfiguration.GetAttributeValue<string>("hil_url");
                inconfig.userName = integrationConfiguration.GetAttributeValue<string>("hil_username");
                inconfig.password = integrationConfiguration.GetAttributeValue<string>("hil_password");
                return inconfig;
            }
            catch (Exception ex)
            {
                throw new Exception("Error : " + ex.Message);
            }
        }
        private static string CallAPI(IntegrationConfiguration integrationConfiguration, string Json, String method)
        {
            WebRequest request = WebRequest.Create(integrationConfiguration.url);
            request.Headers[HttpRequestHeader.Authorization] = "Basic " + Convert.ToBase64String(Encoding.ASCII.GetBytes(integrationConfiguration.userName + ":" + integrationConfiguration.password));
            request.Method = method; //"POST";
            if (!string.IsNullOrEmpty(Json))
            {
                byte[] byteArray = Encoding.UTF8.GetBytes(Json);
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = byteArray.Length;
                Stream dataStream = request.GetRequestStream();
                dataStream.Write(byteArray, 0, byteArray.Length);
                dataStream.Close();
            }
            WebResponse response = request.GetResponse();
            Stream dataStream1 = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream1);
            return reader.ReadToEnd();
        }
        private static string? IsValidMobileNumber(string mobileNumber)
        {
            //Regex pattern to match numbers starting with 6 and having exactly 10 digits
            string pattern = @"^[6-9]\d{9}$";
            if (!string.IsNullOrEmpty(mobileNumber) && Regex.IsMatch(mobileNumber, pattern))
            {
                return mobileNumber;
            }
            return null;
        }
    }
    #region ApiModel's
    public class GetChannelPartnerModel
    {
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public bool IsInitialLoad { get; set; } = false;
        public string CustomerCode { get; set; }
        public string Branch { get; set; }
    }
    public class ChannelPartnerList
    {
        public string Result { get; set; }
        public List<ChannelPartnerModel> Results { get; set; }
    }
    public class ChannelPartnerModel
    {
        public string RetailerCode { get; set; }
        public string RetailerSAPCode { get; set; }
        public string RetailerEmail { get; set; }
        public string RetailerFullName { get; set; }
        public string RetailerFirmName { get; set; }
        public string UserType { get; set; }
        public string UserSubType { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string PINCode { get; set; }
        public string AreaCode { get; set; }
        public string s_MobileNo { get; set; }
        public string PAN { get; set; }
        public bool IsActive { get; set; }
        public string GSTNumber { get; set; }
    }
    public class IntegrationConfiguration
    {
        public string url { get; set; }
        public string userName { get; set; }
        public string password { get; set; }
    }
    #endregion ApiModel's
}
