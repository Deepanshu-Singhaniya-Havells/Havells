using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;

namespace Havells.CRM.GetChannelPartner
{
    public class GetRetailersMaster
    {
        public static void getRetailersMaster(IOrganizationService service, string _syncDatetime)
        {
            try
            {
                Integration intConf = GetPartner.GetIntegration(service, "GetRetailersMaster");
                string uri = "https://p90ci.havells.com:50001/RESTAdapter/MDMService/Core/Partner/GetRetailersMaster";//intConf.uri;
                string authInfo = "D365_HAVELLS:PRDD365@1234";//intConf.userName + ":" + intConf.passWord;
                authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));

                string todate = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') + DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');
                string FromDate = "20230418143015";// (_syncDatetime != string.Empty && _syncDatetime.Trim().Length > 0) ? _syncDatetime : GetPartner.getTimeStamp(service);
                GetRetalierRequest getRetalierRequest = new GetRetalierRequest()
                {
                    FromDate = FromDate,
                    ToDate = todate,
                    IsInitialLoad = false
                };

                //if (_syncDatetime != string.Empty && _syncDatetime.Trim().Length > 0)
                //{
                //    uri = uri + _syncDatetime;
                //}
                //else
                //{
                //    uri = uri + GetPartner.getTimeStamp(service);
                //}
                Console.WriteLine("URL: " + uri);
                string data = JsonConvert.SerializeObject(getRetalierRequest);

                var client = new RestClient(uri);
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);

                request.AddHeader("Authorization", "Basic " + authInfo);
                request.AddHeader("Content-Type", "application/json");

                request.AddParameter("application/json", data, ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);


                //var client = new RestClient(uri);
                //client.Timeout = -1;
                //var request = new RestRequest(Method.POST);
                //request.AddHeader("Authorization", "Basic " + authInfo);
                //Console.WriteLine("Downloading Channel Partner Data: " + DateTime.Now.ToString());
                //IRestResponse response = client.Execute(request);
                GetRetalierResponse rootstring = JsonConvert.DeserializeObject<GetRetalierResponse>(response.Content);
                Console.WriteLine("Downloading Completed of Channel Partner Data: " + DateTime.Now.ToString());
                SyncChannelPartner(rootstring, service);


            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
            }
        }
        private static void SyncChannelPartner(GetRetalierResponse rootstring, IOrganizationService service)
        {
            string _KTOKD = string.Empty;
            Guid IfExists = Guid.Empty;
            string iKUNNR = string.Empty;
            Account iAccount = null;
            EntityReference entRefDistrict = null;
            int iDone = 0;
            int iTotal = rootstring.Results.Count;
            foreach (Result obj in rootstring.Results)
            {
                try
                {
                    iAccount = new Account();
                    if (obj.RetailerCode.ToUpper() != "CCO0440")
                    {
                        IfExists = CheckIfPartnerExists(iKUNNR, service);

                        iAccount.hil_OutWarrantyCustomerSAPCode = obj.RetailerCode.ToUpper();
                        iAccount.Telephone1 = obj.ParentMobileNumber != null ? obj.ParentMobileNumber.ToString() : "";
                        if (obj.RetailerFullName != null && obj.RetailerMobile != null)
                            iAccount.PrimaryContactId = new EntityReference("contatc", retrivePrimaryContact(service, obj.RetailerMobile, obj.RetailerFullName));
                        iAccount.Name = obj.RetailerFirmName != null ? obj.RetailerFirmName : "";
                        Guid UserType = (obj.UserTypeCode == null) ? Guid.Empty : getUserType(obj.UserType, null, obj.UserTypeCode, null, service);
                        Guid userSubType = (obj.UserSubTypeCode == null) ? Guid.Empty : getUserType(obj.UserType, obj.UserSubType, obj.UserTypeCode, obj.UserSubTypeCode, service);
                        if (UserType != Guid.Empty)
                            iAccount["hil_usertype"] = new EntityReference("hil_usertype", UserType);
                        if (userSubType != Guid.Empty)
                            iAccount["hil_usersubtype"] = new EntityReference("hil_usertype", userSubType);

                        //int currentStatus = obj.
                        //iAccount.AccountRatingCode = new OptionSetValue(1);//user current Status

                        iAccount.EMailAddress1 = obj.RetailerEmail;
                        iAccount.Address1_Line1 = obj.Address1;
                        iAccount.Address1_Line2 = obj.Address2;

                        //iAccount.hil_FullAddress = "";//Field Not Found Permanent Address

                        QueryExpression Query = new QueryExpression(hil_businessmapping.EntityLogicalName);
                        Query.ColumnSet = new ColumnSet(true);
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition(new ConditionExpression("hil_stagingarea", ConditionOperator.Equal, obj.AreaCode));
                        Query.Criteria.AddCondition(new ConditionExpression("hil_stagingpin", ConditionOperator.Equal, obj.PINCode));
                        Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        EntityCollection Found = service.RetrieveMultiple(Query);
                        EntityReference salesOfficeRef = new EntityReference();
                        if (Found.Entities.Count > 0)
                        {
                            hil_businessmapping iBusMap = Found.Entities[0].ToEntity<hil_businessmapping>();
                            iAccount.hil_city = iBusMap.hil_city;
                            iAccount.hil_area = iBusMap.hil_area;
                            iAccount.hil_pincode = iBusMap.hil_pincode;
                            iAccount.hil_region = iBusMap.hil_region;
                            iAccount.hil_state = iBusMap.hil_state;
                            iAccount.hil_district = iBusMap.hil_district;
                            iAccount.hil_branch = iBusMap.hil_branch;
                            iAccount.hil_subterritory = iBusMap.hil_subterritory;
                            iAccount.hil_salesoffice = iBusMap.hil_salesoffice;
                            salesOfficeRef = iBusMap.hil_salesoffice;
                            entRefDistrict = iBusMap.hil_district;
                        }
                        if (IfExists == Guid.Empty)
                            service.Create(iAccount);
                        else
                        {
                            iAccount.AccountId = IfExists;
                            service.Update(iAccount);
                        }
                    }
                    else
                    {
                        Console.WriteLine(obj.RetailerCode.ToUpper());
                    }
                    //}
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error :" + ex.Message);
                }
                iDone = iDone + 1;
                Console.WriteLine("Record has been processed :" + iDone + "/" + iTotal);
            }
            Console.WriteLine(" TOTAL COUNT :" + iTotal.ToString());
        }
        public static Guid CheckIfPartnerExists(string KUNNR, IOrganizationService service)
        {
            Guid Partner = new Guid();
            Partner = Guid.Empty;
            QueryExpression Query = new QueryExpression(Account.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_outwarrantycustomersapcode", ConditionOperator.Equal, KUNNR);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                Partner = Found.Entities[0].Id;
            }
            return Partner;
        }
        static Guid retrivePrimaryContact(IOrganizationService service, string mobileNUmber, string fullName)
        {
            QueryExpression qsCType = new QueryExpression("contact");
            qsCType.ColumnSet = new ColumnSet(false);
            qsCType.Criteria = new FilterExpression(LogicalOperator.And);
            qsCType.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, mobileNUmber);
            EntityCollection entCol = service.RetrieveMultiple(qsCType);
            if (entCol.Entities.Count == 1)
                return entCol[0].Id;
            else
            {
                Entity entConsumer = new Entity("contact");
                entConsumer["mobilephone"] = mobileNUmber;
                string[] consumerName = fullName.Split(' ');
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
                    entConsumer["firstname"] = fullName;
                }
                Guid consumerGuId = service.Create(entConsumer);
                return consumerGuId;
            }
        }
        static Guid getUserType(string userTypeName, string userSubTypeName, string userTypeCode, string userSubTypeCode, IOrganizationService service)
        {
            if (userSubTypeCode == "UTC0077")
                Console.Write("d");
            Guid guid = Guid.Empty;
            if (userSubTypeName != null)
            {
                QueryExpression qsCType = new QueryExpression("hil_usertype");
                qsCType.ColumnSet = new ColumnSet(false);
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_code", ConditionOperator.Equal, userTypeCode);
                qsCType.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                qsCType.Criteria.AddCondition("hil_parentusertype", ConditionOperator.Null);
                EntityCollection entCol = service.RetrieveMultiple(qsCType);
                if (entCol.Entities.Count > 0)
                {
                    string fetch = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""true"">
                                  <entity name=""hil_usertype"">
                                    <attribute name=""hil_usertypeid"" />
                                    <attribute name=""hil_code"" />
                                    <attribute name=""createdon"" />
                                    <order attribute=""hil_code"" descending=""false"" />
                                    <filter type=""and"">
                                      <condition attribute=""hil_code"" operator=""eq"" value=""{userSubTypeCode}"" />
                                      <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                                      <condition attribute=""hil_parentusertype"" operator=""eq"" value=""{entCol.Entities[0].Id}"" />
                                    </filter>
                                  </entity>
                                </fetch>";
                    EntityCollection userTypeColl = service.RetrieveMultiple(new FetchExpression(fetch));
                    if (userTypeColl.Entities.Count == 0)
                    {
                        Entity userSubTypeEnt = new Entity("hil_usertype");
                        userSubTypeEnt["hil_name"] = userSubTypeName;
                        userSubTypeEnt["hil_code"] = userSubTypeCode;
                        userSubTypeEnt["hil_parentusertype"] = entCol.Entities[0].ToEntityReference();
                        guid = service.Create(userSubTypeEnt);
                        return guid;
                    }
                    else
                    {
                        guid = userTypeColl[0].Id;
                        return guid;
                    }
                }
                else
                    return Guid.Empty;
            }
            else
            {
                QueryExpression qsCType = new QueryExpression("hil_usertype");
                qsCType.ColumnSet = new ColumnSet(false);
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_code", ConditionOperator.Equal, userTypeCode);
                qsCType.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                qsCType.Criteria.AddCondition("hil_parentusertype", ConditionOperator.Null);
                EntityCollection entCol = service.RetrieveMultiple(qsCType);
                if (entCol.Entities.Count == 0)
                {
                    Entity userSubTypeEnt = new Entity("hil_usertype");
                    userSubTypeEnt["hil_name"] = userTypeName;
                    userSubTypeEnt["hil_code"] = userTypeCode;
                    guid = service.Create(userSubTypeEnt);
                    return guid;
                }
                else
                {
                    guid = entCol[0].Id;
                    return guid;
                }
            }

        }
    }
    // Root myDeserializedClass = JsonConvert.Deserializestring<Root>(myJsonResponse);
    public class Result
    {
        public int RetailerMasterID { get; set; }
        public string RetailerCode { get; set; }
        public string RetailerDMSCode { get; set; }
        public string RetailerSamparkCode { get; set; }
        public string RetailerSAPCode { get; set; }
        public string RetailerFirmName { get; set; }
        public string RetailerFullName { get; set; }
        public string RetailerEmail { get; set; }
        public string RetailerMobile { get; set; }
        public string OwnerGender { get; set; }
        public string RetailerImage { get; set; }
        public string ShopImage { get; set; }
        public string Category { get; set; }
        public string Class { get; set; }
        public string Type { get; set; }
        public string WeeklyOff { get; set; }
        public string TypeOfFirm { get; set; }
        public bool IsEmailConfirmationRequired { get; set; }
        public string TaxRegistration { get; set; }
        public string TIN { get; set; }
        public string TINImagePath { get; set; }
        public string AADHAR { get; set; }
        public string AADHARImagePath { get; set; }
        public string PAN { get; set; }
        public string PANImagePath { get; set; }
        public string KAMCode { get; set; }
        public string GSTNumber { get; set; }
        public string GSTType { get; set; }
        public string GSTNumberImagePath { get; set; }
        public string ParentChildStatus { get; set; }
        public string ParentMobileNumber { get; set; }
        public string AttachedDocumentPath { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public string CountryCode { get; set; }
        public string CountryName { get; set; }
        public string StateCode { get; set; }
        public string StateName { get; set; }
        public string GSTStateCode { get; set; }
        public string GSTStateName { get; set; }
        public string DistrictCode { get; set; }
        public string DistrictName { get; set; }
        public string CityCode { get; set; }
        public string CityName { get; set; }
        public string AreaCode { get; set; }
        public string AreaName { get; set; }
        public string PINCode { get; set; }
        public string ServiceBranchCode { get; set; }
        public string ServiceBranchName { get; set; }
        public string SalesOfficeCode { get; set; }
        public string SalesOfficeName { get; set; }
        public string ResidenceAddress1 { get; set; }
        public string ResidenceAddress2 { get; set; }
        public DateTime DOB { get; set; }
        public string Lattitude { get; set; }
        public string Longitude { get; set; }
        public string RetailerType { get; set; }
        public string MainContactName { get; set; }
        public string MainContactMobileNo { get; set; }
        public string MainContactEmail { get; set; }
        public string AnniversaryDate { get; set; }
        public string SpouseName { get; set; }
        public string BankAccountNumber { get; set; }
        public string IFSCCode { get; set; }
        public string BankName { get; set; }
        public string BankAddress { get; set; }
        public string BankAccountName { get; set; }
        public string ChequeCancelImagePath { get; set; }
        public string Document1 { get; set; }
        public string Document2 { get; set; }
        public string Document3 { get; set; }
        public string Document4 { get; set; }
        public string Document5 { get; set; }
        public string CreatorSource { get; set; }
        public string Divisions { get; set; }
        public bool IsBulkUpload { get; set; }
        public string LastSyncDateDMS { get; set; }
        public DateTime RetailerRegistrationDate { get; set; }
        public bool IsActive { get; set; }
        public DateTime CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
        public string RetailerChannel { get; set; }
        public string RetailerClass { get; set; }
        public string RetailerGroup { get; set; }
        public string UserType { get; set; }
        public string IsTCS { get; set; }
        public string UserSubType { get; set; }
        public string UserCode { get; set; }
        public string Education { get; set; }
        public string RefSapCode { get; set; }
        public string sendToService { get; set; }
        public string IsVerified { get; set; }
        public string ManagerName { get; set; }
        public string ManagerMobile { get; set; }
        public string ManagerEmail { get; set; }
        public string IMEI1 { get; set; }
        public string IMEI2 { get; set; }
        public string EmployeeCode { get; set; }
        public bool IsAllowLimit { get; set; }
        public string IsCNActive { get; set; }
        public string ReferenceId { get; set; }
        public DateTime Eff_FromDate { get; set; }
        public DateTime Eff_ToDate { get; set; }
        public string CreatedSource { get; set; }
        public string ModifySource { get; set; }
        public string ModifyBy { get; set; }
        public DateTime? ModifyDate { get; set; }
        public string UserCatCode { get; set; }
        public string UserTypeCode { get; set; }
        public string UserSubTypeCode { get; set; }
        public string FirmCode { get; set; }
        public string DrivingLicence { get; set; }
        public string VoterId { get; set; }
        public string IsTDS { get; set; }
        public string IsITRFiled { get; set; }
    }
    public class GetRetalierResponse
    {
        public string Result { get; set; }
        public List<Result> Results { get; set; }
        public bool Success { get; set; }
        public string Message { get; set; }
        public string ErrorCode { get; set; }
    }
    // Root myDeserializedClass = JsonConvert.Deserializestring<Root>(myJsonResponse);
    public class GetRetalierRequest
    {
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public bool IsInitialLoad { get; set; }
    }


}
