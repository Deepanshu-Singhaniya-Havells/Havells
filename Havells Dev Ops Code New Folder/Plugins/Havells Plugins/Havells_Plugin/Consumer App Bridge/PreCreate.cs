#region Definitions
/*
Error Codes
    Error Code : 1 - Customer Not Registered
    Error Code : 2 - Duplicate Email Ids Exist
    Error Code : 3 - Operation Code Blank
    Error Code : 4 - Missing Required Information
    Error Code : 5 - Invalid Operation code
    Error Code : 6 - Invalid Password
    Error Code : 7 - Invalid Email
    Error Code : 8 - Duplicate Products Exist
    Error Code : 9 - Product Category/Product/Serial Number not Found
    Error Code : 10 - CRM Error
    Error Code : 11 - Record Not Found
    Error Code : 12 - Dummy Customer Not Found
    Error Code : 13 - EmailId Already Exists
    Error Code : 14 - Mobile Number Already

Status Codes
    Status Code : false - Failure
    Status Code : true - Success
     
Function Codes
    Function Code : 1 - Customer Registration
    Function Code : 2 - Customer Login
    Function Code : 3 - Forgot Password
    Function Code : 4 - Change Password
    Function Code : 5 - Product Registration
    Function Code : 6 - Create Work Order Demo
    Function Code : 7 - Create Work Order Service
    Function Code : 8 - Profile Update
*/
#endregion

using System;
using System.IO;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Security.Cryptography;

namespace Havells_Plugin.Consumer_App_Bridge
{
    #region Validate Class
    public class Validate
    {
        public bool IfValidated;
        public bool StatusCode;
        public int FunctionCode;
        public string OTP;
        public string ContactGuID;
        public string ExceptionCode;
        public string ExceptionDesc;
        public string NewPassword;
        public string WOUniqueID;
        public string MobileNumber;
        public Guid CustomerAssetId;
        public string WOUniqueReferenceNumber;
        public string FirstName;
        public string LastName;
        public string PinCode;
        public bool ExistingUser;
    }
    #endregion
    #region Shared Variable Class
    public class SharedVariables
    {
        public EntityReference ProdID = new EntityReference(Product.EntityLogicalName);
        public EntityReference ProdGrp = new EntityReference(Product.EntityLogicalName);
        public string PrdtSlNo = string.Empty;
    }
    #endregion
    public class PreCreate : IPlugin
    {
        #region Execute Method
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_consumerappbridge.EntityLogicalName
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    hil_consumerappbridge Brdg = entity.ToEntity<hil_consumerappbridge>();
                    if( Brdg.Attributes.Contains("hil_productsubcategorymap") && 
                       Brdg.hil_ServiceRequestType != null && Brdg.hil_FirstName != null && 
                       Brdg.hil_LastName != null && Brdg.hil_MobileNumber != null && Brdg.hil_EmailId != null)
                    {
                        CreateWOForPortalDemo(service, Brdg);
                    }
                    else
                    {
                        InitiateExecution(service, Brdg, tracingService);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ConsumerAppBridge.ConsumerAppBridge_PreValidation" + ex.Message);
            }
            #endregion
        }
        #endregion
        #region WorkOrder For Portal
        public static void CreateWOForPortalDemo(IOrganizationService service, hil_consumerappbridge Brdg)
        {
            Guid ContId = new Guid();
            ContId = CheckIfUserExists(service, Brdg);
            Guid CallSubType = new Guid();
            Guid Nature = new Guid();
            Guid ServiceAccount = Helper.GetGuidbyName(Account.EntityLogicalName, "name", "Dummy Customer", service);
            Guid PriceList = Helper.GetGuidbyName(PriceLevel.EntityLogicalName, "name", "Default Price List", service);
            msdyn_workorder _WrkOd = new msdyn_workorder();
            EntityReference iCatSubCat = (EntityReference)Brdg["hil_productsubcategorymap"];
            _WrkOd.hil_ProductCatSubCatMapping = iCatSubCat;
            hil_stagingdivisonmaterialgroupmapping iMap = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, iCatSubCat.Id, new ColumnSet(true));
            if (Brdg.hil_ServiceRequestType.Value == 1)//Installation
            {
                CallSubType = Helper.GetGuidbyName(hil_callsubtype.EntityLogicalName, "hil_name", "Installation", service);
                Nature = GetNature(service, "Installation", iMap.hil_ProductSubCategoryMG);
            }
            else if(Brdg.hil_ServiceRequestType.Value == 2)//Demo
            {
                CallSubType = Helper.GetGuidbyName(hil_callsubtype.EntityLogicalName, "hil_name", "Demo", service);
                Nature = GetNature(service, "Demo", iMap.hil_ProductSubCategoryMG);
            }
            _WrkOd.hil_CustomerRef = new EntityReference(Contact.EntityLogicalName, ContId);
            _WrkOd.hil_mobilenumber = Brdg.hil_MobileNumber;
            _WrkOd.msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
            _WrkOd.msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
            if (CallSubType != Guid.Empty)
            {
                _WrkOd.hil_CallSubType = new EntityReference(hil_callsubtype.EntityLogicalName, CallSubType);
            }
            if (PriceList != Guid.Empty)
            {
                _WrkOd.msdyn_PriceList = new EntityReference(PriceLevel.EntityLogicalName, PriceList);
            }
            if (Nature != Guid.Empty)
            {
                _WrkOd.hil_natureofcomplaint = new EntityReference(hil_natureofcomplaint.EntityLogicalName, Nature);
            }
            if (Brdg.hil_Description != null)
            {
                _WrkOd["hil_customercomplaintdescription"] = Brdg.hil_Description;
            }
        EntityReference CallType = new EntityReference(hil_calltype.EntityLogicalName);
            hil_callsubtype Call = (hil_callsubtype)service.Retrieve(hil_callsubtype.EntityLogicalName, CallSubType, new ColumnSet(true));
            if (Call.hil_CallType != null)
            {
                CallType = Call.hil_CallType;
                _WrkOd.hil_CallType = new EntityReference(hil_calltype.EntityLogicalName, CallType.Id);
            }
            if(iMap.hil_ProductCategoryDivision != null)
                _WrkOd.hil_Productcategory = iMap.hil_ProductCategoryDivision;
            if(iMap.hil_ProductSubCategoryMG != null)
               _WrkOd.hil_ProductSubcategory = iMap.hil_ProductSubCategoryMG;
            _WrkOd.hil_quantity = 1;
            _WrkOd.hil_SourceofJob = new OptionSetValue(3);
            QueryExpression Query = new QueryExpression(hil_address.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, ContId);
            Query.Criteria.AddCondition("hil_addresstype", ConditionOperator.Equal, 1);
            Query.AddOrder("createdon", OrderType.Descending);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count > 0)
            {
                hil_address iAdd = Found.Entities[0].ToEntity<hil_address>();
                _WrkOd.hil_Address = new EntityReference(hil_address.EntityLogicalName, iAdd.Id);
                if(iAdd.hil_PinCode != null)
                _WrkOd.hil_pincode = iAdd.hil_PinCode;
                if(iAdd.hil_CIty != null)
                _WrkOd.hil_City = iAdd.hil_CIty;
                if(iAdd.hil_State != null)
                _WrkOd.hil_state = iAdd.hil_State;
                if (iAdd.hil_SalesOffice != null)
                {
                    _WrkOd["hil_salesoffice"] = iAdd.hil_SalesOffice;
                } 
                else if(iAdd.hil_SalesOffice == null && iAdd.hil_PinCode != null)
                {
                    hil_businessmapping BusMap = GetBusinessMapping(service, iAdd.hil_PinCode.Id);
                    if (BusMap.hil_salesoffice != null)
                        _WrkOd["hil_salesoffice"] = BusMap.hil_salesoffice;
                }
            }
            Guid _WoID = service.Create(_WrkOd);

        }
        #region Find Contact
        public static Guid CheckIfUserExists(IOrganizationService service, hil_consumerappbridge Brdg)
        {
            Guid ContId = new Guid();
            QueryExpression Query = new QueryExpression();
            Query.EntityName = Contact.EntityLogicalName;
            ColumnSet Col = new ColumnSet(true);
            Query.ColumnSet = Col;
            Query.Criteria = new FilterExpression(LogicalOperator.Or);
            Query.Criteria.AddCondition("emailaddress1", ConditionOperator.Equal, Brdg.hil_EmailId);
            Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, Brdg.hil_MobileNumber);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count >= 1)
            {
                foreach(Contact Ct in Found.Entities)
                {
                    ContId = Ct.ContactId.Value;
                }
            }
            else
            {
                if(Brdg.hil_AddressLine1 != null && Brdg.Attributes.Contains("hil_businessgeo"))
                {
                    Contact Ct = new Contact();
                    Ct.EMailAddress1 = Brdg.hil_EmailId;
                    Ct.MobilePhone = Brdg.hil_MobileNumber;
                    Ct.FirstName = Brdg.hil_FirstName;
                    Ct.LastName = Brdg.hil_LastName;
                    ContId = service.Create(Ct);
                    hil_address Add = new hil_address();
                    Add.hil_AddressType = new OptionSetValue(1);
                    Add.hil_Street1 = Brdg.hil_AddressLine1;
                    Add.hil_Street2 = Brdg.hil_AddressLine2;
                    EntityReference BusinessGeo = (EntityReference)Brdg["hil_businessgeo"];
                    Add.hil_BusinessGeo = BusinessGeo;
                    hil_businessmapping iMap = (hil_businessmapping)service.Retrieve(hil_businessmapping.EntityLogicalName, BusinessGeo.Id, new ColumnSet(true));
                    Add.hil_CIty = iMap.hil_city;
                    Add.hil_State = iMap.hil_state;
                    Add.hil_PinCode = iMap.hil_pincode;
                    Add.hil_Region = iMap.hil_region;
                    Add.hil_Area = iMap.hil_area;
                    Add.hil_Branch = iMap.hil_branch;
                    Add.hil_SalesOffice = iMap.hil_salesoffice;
                    Add.hil_District = iMap.hil_district;
                    Add.hil_SubTerritory = iMap.hil_subterritory;
                    Add.hil_Customer = new EntityReference(Contact.EntityLogicalName, ContId);
                    service.Create(Add);
                }
                else
                {
                    throw new InvalidPluginExecutionException("Please Enter Address Details");
                }
            }
            return (ContId);
        }
        #endregion
        #endregion
        #region Identify Operation
        public static void InitiateExecution(IOrganizationService service, hil_consumerappbridge Brdg, ITracingService Tracing)
        {
            try
            {
                Validate VD = new Validate();
                string Method = string.Empty;
                if (Brdg.hil_MethodName != null)
                {
                    Method = (string)Brdg.hil_MethodName;
                    switch (Method)
                    {
                        case "Create":
                            //VD.FunctionCode = 1;
                            if ((Brdg.hil_LastName != null) && (Brdg.hil_Password != null) && (Brdg.hil_UserName != null))
                            {
                                VD.FunctionCode = 1;
                                VD = CreateContact(service, Brdg, VD.FunctionCode);
                                VD.FunctionCode = 1;
                            }
                            else
                            {
                                VD.StatusCode = false;
                                VD.ExceptionCode = "4";
                                VD.ExceptionDesc = "Missing Required Information";
                                VD.ContactGuID = "Not Created";
                                VD.FunctionCode = 1;
                            }
                            break;
                        case "CustomerLogin":
                            //VD.FunctionCode = 2;
                            if ((Brdg.hil_UserName != null) && (Brdg.hil_Password != null))
                            {
                                VD = ValidateCustomer(service, Brdg);
                                VD.FunctionCode = 2;
                            }
                            else
                            {
                                VD.StatusCode = false;
                                VD.ExceptionCode = "4";
                                VD.ExceptionDesc = "Missing Required Information";
                                VD.FunctionCode = 2;
                                VD.IfValidated = false;
                            }
                            break;
                        case "ForgotPassword":
                            //VD.FunctionCode = 3;
                            if (Brdg.hil_UserName != null)
                            {
                                VD = ForgotPassword(service, Brdg);
                                VD.FunctionCode = 3;
                            }
                            else
                            {
                                VD.StatusCode = false;
                                VD.ExceptionCode = "4";
                                VD.ExceptionDesc = "Missing Required Information";
                                VD.FunctionCode = 3;
                                VD.OTP = "XX";
                            }
                            break;
                        case "ChangePassword":
                            //VD.FunctionCode = 4;
                            if ((Brdg.hil_UserName != null) && (Brdg.hil_Password != null))
                            {
                                VD = ChangePasswordValidate(service, Brdg);
                                VD.FunctionCode = 4;
                            }
                            else
                            {
                                VD.StatusCode = false;
                                VD.ExceptionCode = "4";
                                VD.ExceptionDesc = "Missing Required Information";
                                VD.FunctionCode = 4;
                            }
                            break;
                        case "ProductRegistration":
                            //VD.FunctionCode = 5;
                            if ((Brdg.hil_ContactGuId != null) && (Brdg.hil_PinCode != null) && 
                                (Brdg.hil_InvDate != null) && (Brdg.hil_InvoiceName != null) && 
                                (Brdg.hil_Base64String != null) && (Brdg.hil_FileType != null) && 
                                (Brdg.hil_ProductItemCode != null) && Brdg.hil_ShippingState != null)
                            {
                                if (Brdg.hil_ProductSNo != null)
                                {
                                    Tracing.Trace(Convert.ToString(DateTime.Now));
                                    VD.FunctionCode = 5;
                                    VD = CreateCustomerAsset(service, new Guid(Brdg.hil_ContactGuId), Convert.ToDateTime(Brdg.hil_InvDate),
                                        Brdg.hil_InvoiceName, Brdg.hil_ProductSNo, VD.FunctionCode, Brdg.hil_Base64String,
                                        Convert.ToInt16(Brdg.hil_FileType), Brdg.hil_ProductItemCode, Brdg.hil_ShippingCountry, Brdg.hil_ShippingState, Brdg);
                                    Tracing.Trace(Convert.ToString(DateTime.Now));
                                    VD.FunctionCode = 5;
                                    Tracing.Trace(Convert.ToString(DateTime.Now));
                                }
                                else
                                {
                                    Tracing.Trace(Convert.ToString(DateTime.Now));
                                    VD.FunctionCode = 5;
                                    VD = CreateCustomerAsset(service, new Guid(Brdg.hil_ContactGuId), Convert.ToDateTime(Brdg.hil_InvDate),
                                        Brdg.hil_InvoiceName, "", VD.FunctionCode, Brdg.hil_Base64String,
                                        Convert.ToInt16(Brdg.hil_FileType), Brdg.hil_ProductItemCode, Brdg.hil_ShippingCountry, Brdg.hil_ShippingState, Brdg);
                                    Tracing.Trace(Convert.ToString(DateTime.Now));
                                    VD.FunctionCode = 5;
                                    Tracing.Trace(Convert.ToString(DateTime.Now));
                                }
                            }
                            else
                            {
                                VD.StatusCode = false;
                                VD.ExceptionCode = "5";
                                VD.ExceptionDesc = "Missing Required Information";
                                VD.FunctionCode = 5;
                            }
                            break;
                        case "CreateWODemo":
                            VD.FunctionCode = 6;
                            Guid ServiceAccount = Helper.GetGuidbyName(Account.EntityLogicalName, "name", "Dummy Customer", service);
                            if(ServiceAccount != Guid.Empty)
                            {
                                if (Brdg.hil_ContactGuId != null && Brdg.hil_IncidentType != null && Brdg.hil_ProductItemCode != null)
                                {
                                    VD.FunctionCode = 6;
                                    string NatureOfComplaintDesc = string.Empty;
                                    VD = CreateWorkOrderForDemo(service, Brdg.hil_ContactGuId, ServiceAccount, Brdg.hil_ProductItemCode, Brdg.hil_IncidentType, Brdg.hil_Description);
                                    //VD = IfSerialNumberExists(service, Brdg.hil_ProductSNo, Brdg.hil_ProductItemCode, Brdg.hil_ContactGuId);
                                    VD.FunctionCode = 6;
                                }
                                else if (Brdg.hil_FirstName != null && Brdg.hil_LastName != null && Brdg.hil_MobileNumber != null && Brdg.hil_UserName != null && Brdg.hil_ShippingPincode != null && Brdg.hil_IncidentType != null && Brdg.hil_ProductItemCode != null)
                                {
                                    VD.FunctionCode = 6;
                                    Guid ContGuid = GetThisCustomerIfExists(service, Brdg.hil_MobileNumber, Brdg.hil_UserName);
                                    if(ContGuid == Guid.Empty)
                                    {
                                        VD = CreateContact(service, Brdg, VD.FunctionCode);
                                        if (VD.StatusCode == true)
                                        {
                                            string NatureOfComplaintDesc = string.Empty;
                                            VD = CreateWorkOrderForDemo(service, VD.ContactGuID, ServiceAccount, Brdg.hil_ProductItemCode, Brdg.hil_IncidentType, Brdg.hil_Description);
                                            VD.FunctionCode = 6;
                                        }
                                    }
                                    else
                                    {
                                        string NatureOfComplaintDesc = string.Empty;
                                        VD = CreateWorkOrderForDemo(service, Convert.ToString(ContGuid), ServiceAccount, Brdg.hil_ProductItemCode, Brdg.hil_IncidentType, Brdg.hil_Description);
                                        VD.FunctionCode = 6;
                                    }
                                    VD.FunctionCode = 6;
                                }
                                else
                                {
                                    VD.StatusCode = false;
                                    VD.ExceptionCode = "5";
                                    VD.ExceptionDesc = "Missing Required Information";
                                    VD.WOUniqueID = "Not Created";
                                    VD.FunctionCode = 6;
                                    VD.ContactGuID = "Not Created";
                                }
                            }
                        
                            else
                            {
                                VD.FunctionCode = 7;
                                VD.StatusCode = false;
                                VD.ExceptionCode = "12";
                                VD.ExceptionDesc = "Dummy Customer not Found";
                                VD.ContactGuID = "NOT FOUND";
                                VD.WOUniqueID = "Not Created";
                            }
                            break;
                        case "CreateWOService":
                            VD.FunctionCode = 7;
                            if (Brdg.hil_ContactGuId != null && Brdg.hil_ProductItemCode != null && Brdg.hil_IncidentType != null)
                            {
                                string NatureOfComplaintDesc = string.Empty;
                                if(Brdg.hil_NatureOfComplaint != null)
                                {
                                    NatureOfComplaintDesc = Brdg.hil_NatureOfComplaint;
                                }
                                VD = CreateWorkOrderForService(service, Brdg.hil_ContactGuId, NatureOfComplaintDesc, Brdg.hil_IncidentType, Brdg.hil_ProductItemCode, Brdg.hil_Description, Brdg.hil_Country);
                                VD.FunctionCode = 7;
                            }
                            else
                            {
                                VD.StatusCode = false;
                                VD.ExceptionCode = "5";
                                VD.ExceptionDesc = "Missing Required Information";
                                VD.WOUniqueID = "Not Created";
                                VD.FunctionCode = 7;
                                VD.ContactGuID = "Not Created";
                            }
                            break;
                        case "ProfileUpdate":
                            if (Brdg.hil_ContactGuId != null)
                            {
                                VD.FunctionCode = 8;
                                VD = ContactProfileUpdate(service, Brdg);
                                VD.FunctionCode = 8;
                            }
                            else
                            {
                                VD.StatusCode = false;
                                VD.ExceptionCode = "5";
                                VD.ExceptionDesc = "Missing Required Information";
                                VD.ContactGuID = "Not Found";
                                VD.FunctionCode = 8;
                            }
                            break;
                        default:
                            VD.StatusCode = false;
                            VD.ExceptionCode = "5";
                            VD.ExceptionDesc = "Invalid Operation code";
                            break;
                    }
                }
                else
                {
                    VD.StatusCode = false;
                    VD.ExceptionCode = "3";
                    VD.ExceptionDesc = "Operation Code Blank";
                }
                UpdateBridgeRecord(service, VD, Brdg, Tracing);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Exception Occured at :" + ex.Message);
            }
        }
        #endregion
        #region Create Contact
        public static Validate CreateContact(IOrganizationService service, hil_consumerappbridge Brdg, int FunctionCode)
        {
            Validate VD = new Validate();
            Contact Cont = new Contact();
            VD = CheckThisUserName(service, Brdg.hil_UserName, FunctionCode);
            if(VD.ExistingUser == false)
            {
                VD = CheckThisMobileNumber(service, Brdg.hil_MobileNumber, FunctionCode);
                if (VD.ExistingUser == false)
                {
                    if (Brdg.hil_LastName != null)
                    {
                        Cont.LastName = Brdg.hil_LastName;
                    }
                    if (Brdg.hil_UserName != null)
                    {
                        Cont.EMailAddress1 = Brdg.hil_UserName;
                    }
                    if (Brdg.hil_Password != null)
                    {
                        //string _DecryptedPwd = string.Empty;
                       // _DecryptedPwd = DecryptThisPassword(service, Brdg.hil_Password);
                        Cont.hil_Password = Brdg.hil_Password;
                    }
                    if(Brdg.Attributes.Contains("hil_salutationcode"))
                    {
                        int SalCode = (int)Brdg["hil_salutationcode"];
                        OptionSetValue SalOp = new OptionSetValue(SalCode);
                        Cont.hil_Salutation = SalOp;
                    }
                    if (Brdg.hil_FirstName != null)
                    {
                        Cont.FirstName = Brdg.hil_FirstName;
                    }
                    if (Brdg.hil_MobileNumber != null)
                    {
                        Cont.MobilePhone = Brdg.hil_MobileNumber;
                    }
                    if (Brdg.hil_AddressLine1 != null)
                    {
                        Cont.Address1_Line1 = Brdg.hil_AddressLine1;
                    }
                    if (Brdg.hil_AddressLine2 != null)
                    {
                        Cont.Address1_Line2 = Brdg.hil_AddressLine2;
                    }
                    if (Brdg.hil_AddressLine3 != null)
                    {
                        Cont.Address1_Line3 = Brdg.hil_AddressLine3;
                    }
                    if (Brdg.hil_City != null)
                    {
                        Cont.hil_city = new EntityReference(hil_city.EntityLogicalName, new Guid(Brdg.hil_City));
                    }
                    if (Brdg.hil_PinCode != null)
                    {
                        Cont.hil_pincode = new EntityReference(hil_pincode.EntityLogicalName, new Guid(Brdg.hil_PinCode));
                    }
                    if (Brdg.hil_Country != null)
                    {
                        Cont.hil_Country = new EntityReference(hil_country.EntityLogicalName, new Guid(Brdg.hil_Country));
                    }
                    if (Brdg.hil_State != null)
                    {
                        Cont.hil_state = new EntityReference(hil_state.EntityLogicalName, new Guid(Brdg.hil_State));
                    }
                    Guid ContactId = service.Create(Cont);
                    if (ContactId != Guid.Empty)
                    {
                        VD.ContactGuID = ContactId.ToString();
                        VD.StatusCode = true;
                        VD.ExceptionCode = "XX";
                        VD.ExceptionDesc = "XX";
                        VD.FunctionCode = FunctionCode;
                    }
                }
            }
            return (VD);
        }
        public static Validate CheckThisUserName(IOrganizationService service, string Email, int FunctionCode)
        {
            Validate VD = new Validate();
            QueryByAttribute Query = new QueryByAttribute();
            Query.EntityName = Contact.EntityLogicalName;
            Query.AddAttributeValue("emailaddress1", Email);
            ColumnSet Col = new ColumnSet("contactid");
            Query.ColumnSet = Col;
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count >= 1)
            {
                VD.FunctionCode = FunctionCode;
                VD.ExistingUser = true;
                VD.StatusCode = false;
                VD.ExceptionCode = "13";
                VD.ExceptionDesc = "Email Id Already Exists";
                VD.ContactGuID = "Not Created";
            }
            else
            {
                VD.ExistingUser = false;
                VD.StatusCode = true;
                VD.FunctionCode = FunctionCode;
            }
            return VD;
        }
        public static Validate CheckThisMobileNumber(IOrganizationService service, string MobNum, int FunctionCode)
        {
            Validate VD = new Validate();
            QueryByAttribute Query = new QueryByAttribute();
            Query.EntityName = Contact.EntityLogicalName;
            Query.AddAttributeValue("mobilephone", MobNum);
            ColumnSet Col = new ColumnSet("contactid");
            Query.ColumnSet = Col;
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count >= 1)
            {
                VD.FunctionCode = FunctionCode;
                VD.ExistingUser = true;
                VD.StatusCode = false;
                VD.ExceptionCode = "14";
                VD.ExceptionDesc = "Mobile Number Already Exists";
                VD.ContactGuID = "Not Created";
            }
            else
            {
                VD.ExistingUser = false;
                VD.StatusCode = true;
                VD.FunctionCode = FunctionCode;
            }
            return VD;
        }
        #endregion
        #region Validate Customer
        public static Validate ValidateCustomer(IOrganizationService service, hil_consumerappbridge Brdg)
        {
            Validate VD = new Validate();
            string _Decrpypted = string.Empty;
            //_Decrpypted = DecryptThisPassword(service, Brdg.hil_Password);
            QueryByAttribute Query = new QueryByAttribute(Contact.EntityLogicalName);
            ColumnSet Col = new ColumnSet("emailaddress1", "hil_password", "firstname", "lastname", "mobilephone");
            Query.AddAttributeValue("emailaddress1", Brdg.hil_UserName);
            Query.ColumnSet = Col;
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count == 1)
            {
                foreach (Contact Ct in Found.Entities)
                {
                    if (Ct.hil_Password == Brdg.hil_Password)
                    {
                        VD.IfValidated = true;
                        VD.ContactGuID = Ct.Id.ToString();
                        VD.FirstName = Ct.FirstName;
                        VD.LastName = Ct.LastName;
                        VD.MobileNumber = Ct.MobilePhone;
                        VD.PinCode = GetPinCode(service, Ct.Id);
                        VD.StatusCode = true; //Success
                    }
                    else
                    {
                        VD.IfValidated = false;
                        VD.StatusCode = true;
                        VD.MobileNumber = "Not Found";
                        VD.ContactGuID = "Not Found";
                        VD.LastName = "Not Found";
                        VD.FirstName = "Not Found";
                        VD.PinCode = "Not Found";
                        //VD.ExceptionCode = "6";
                        //VD.ExceptionDesc = "Invalid Password";
                    }
                }
            }
            else if (Found.Entities.Count > 1)
            {
                VD.StatusCode = false;
                VD.IfValidated = false;
                VD.ExceptionCode = "2";
                VD.ExceptionDesc = "Duplicate Email Ids Exist";
            }
            else
            {
                VD.IfValidated = false;
                VD.StatusCode = true;
                VD.ContactGuID = "Not Found";
                //VD.ExceptionCode = "1";
                //VD.ExceptionDesc = "Customer Not Registered";
            }
            return VD;
        }
        #endregion
        #region Forgot Password
        public static Validate ForgotPassword(IOrganizationService service, hil_consumerappbridge Brdg)
        {
            Validate VD = new Validate();
            QueryByAttribute Query = new QueryByAttribute(Contact.EntityLogicalName);
            ColumnSet Col = new ColumnSet(false);
            Query.AddAttributeValue("emailaddress1", Brdg.hil_UserName);
            Query.ColumnSet = Col;
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count == 1)
            {
                foreach (Contact Ct in Found.Entities)
                {
                    int iOTPLength = 6;
                    string[] saAllowedCharacters = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
                    VD.OTP = GenerateRandomOTP(iOTPLength, saAllowedCharacters);
                    VD.StatusCode = true;
                    VD.ContactGuID = Ct.Id.ToString();
                    Ct.hil_OTP = VD.OTP;
                    service.Update(Ct);
                }
            }
            else if (Found.Entities.Count > 1)
            {
                VD.StatusCode = false;
                VD.OTP = "XXXX";
                VD.ExceptionCode = "2";
                VD.ExceptionDesc = "Duplicate Email Ids Exist";
            }
            else
            {
                QueryByAttribute Query1 = new QueryByAttribute(Contact.EntityLogicalName);
                ColumnSet Col1 = new ColumnSet(false);
                Query1.AddAttributeValue("mobilephone", Brdg.hil_UserName);
                Query1.ColumnSet = Col1;
                EntityCollection Found1 = service.RetrieveMultiple(Query1);
                if (Found1.Entities.Count == 1)
                {
                    foreach (Contact Ct in Found1.Entities)
                    {
                        int iOTPLength = 6;
                        string[] saAllowedCharacters = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
                        VD.OTP = GenerateRandomOTP(iOTPLength, saAllowedCharacters);
                        VD.StatusCode = true;
                        VD.ContactGuID = Ct.Id.ToString();
                        Ct.hil_OTP = VD.OTP;
                        service.Update(Ct);
                    }
                }
                else if (Found.Entities.Count > 1)
                {
                    VD.StatusCode = false;
                    VD.OTP = "XXXX";
                    VD.ExceptionCode = "2";
                    VD.ExceptionDesc = "Duplicate Email / PhoneNumber Ids Exist";
                }
                else
                {
                    VD.StatusCode = false;
                    VD.OTP = "XXXX";
                    VD.ExceptionCode = "1";
                    VD.ExceptionDesc = "Consumer Not Registered";
                }
            }
            return (VD);
        }
        #endregion
        #region Change Password
        public static Validate ChangePasswordValidate(IOrganizationService service, hil_consumerappbridge BrDg)
        {
            Validate VD = new Validate();
            string _Decrypted = string.Empty;
            QueryByAttribute Query = new QueryByAttribute(Contact.EntityLogicalName);
            ColumnSet Col = new ColumnSet(false);
            Query.AddAttributeValue("emailaddress1", BrDg.hil_UserName);
            Query.ColumnSet = Col;
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count == 1)
            {
                foreach (Contact Ct in Found.Entities)
                {
                    // _Decrypted = DecryptThisPassword(service, BrDg.hil_Password);
                    VD.NewPassword = BrDg.hil_Password;
                    Ct.hil_Password = BrDg.hil_Password;
                    service.Update(Ct);
                    VD.StatusCode = true;
                    VD.ContactGuID = Ct.Id.ToString();
                }
            }
            else if (Found.Entities.Count > 1)
            {
                VD.StatusCode = false;
                VD.ExceptionCode = "2";
                VD.ExceptionDesc = "Duplicate Email Ids Exist";
            }
            else
            {
                QueryByAttribute Query1 = new QueryByAttribute(Contact.EntityLogicalName);
                ColumnSet Col1 = new ColumnSet(false);
                Query1.AddAttributeValue("mobilephone", BrDg.hil_UserName);
                Query1.ColumnSet = Col1;
                EntityCollection Found1 = service.RetrieveMultiple(Query1);
                if (Found1.Entities.Count == 1)
                {
                    foreach (Contact Ct in Found1.Entities)
                    {
                        // _Decrypted = DecryptThisPassword(service, BrDg.hil_Password);
                        VD.NewPassword = BrDg.hil_Password;
                        Ct.hil_Password = BrDg.hil_Password;
                        service.Update(Ct);
                        VD.StatusCode = true;
                        VD.ContactGuID = Ct.Id.ToString();
                    }
                }
                else if (Found.Entities.Count > 1)
                {
                    VD.StatusCode = false;
                    VD.ExceptionCode = "2";
                    VD.ExceptionDesc = "Duplicate Mobile Number Exist";
                }
            }
            return (VD);
        }
        #endregion
        #region Generate OTP
        public static string GenerateRandomOTP(int iOTPLength, string[] saAllowedCharacters)
        {
            string sOTP = String.Empty;
            string sTempChars = String.Empty;
            Random rand = new Random();
            for (int i = 0; i < iOTPLength; i++)
            {
                int p = rand.Next(0, saAllowedCharacters.Length);
                sTempChars = saAllowedCharacters[rand.Next(0, saAllowedCharacters.Length)];
                sOTP += sTempChars;
            }
            return sOTP;
        }
        #endregion
        #region Create Customer Asset
        public static Validate CreateCustomerAsset(IOrganizationService service, Guid ContGuId, DateTime InvoiceDate, string InvoiceName,
            string ProductSNo, int FncCode, string Base64String, int FileType, string DivisionMGStage, string Model, string DPin,  hil_consumerappbridge Brdg)
        {
            Validate VD = new Validate();
            hil_stagingdivisonmaterialgroupmapping Stage = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, new Guid(DivisionMGStage), new ColumnSet("hil_productcategorydivision", "hil_productsubcategorymg"));
            Guid DummyAc = Helper.GetGuidbyName(Account.EntityLogicalName, "name", "Dummy Customer", service);
            msdyn_customerasset _CstA = new msdyn_customerasset();
            if (ProductSNo != null)
            {
                _CstA.msdyn_name = ProductSNo;
            }
            if (ContGuId != Guid.Empty)
            {
                _CstA.hil_Customer = new EntityReference(Contact.EntityLogicalName, ContGuId);
            }
            if(Model != null)
            {
                _CstA.msdyn_Product = new EntityReference(Product.EntityLogicalName, new Guid(Model));
            }
            if(DPin != null)
            {
                QueryExpression Qry = new QueryExpression(hil_pincode.EntityLogicalName);
                //Qry.EntityName = hil_pincode.EntityLogicalName;
                ColumnSet Col = new ColumnSet("hil_name", "hil_pincodeid", "hil_city", "hil_state");
                Qry.ColumnSet = Col;
                Qry.Criteria = new FilterExpression(LogicalOperator.And);
                Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                Qry.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, DPin));
                EntityCollection Colec = service.RetrieveMultiple(Qry);
                if (Colec.Entities.Count >= 1)
                {
                    hil_pincode iPin = Colec.Entities[0].ToEntity<hil_pincode>();
                    EntityReference State = new EntityReference();
                    EntityReference City = new EntityReference();
                    QueryExpression Qry1 = new QueryExpression(hil_businessmapping.EntityLogicalName);
                    ColumnSet Col1 = new ColumnSet(false);
                    Qry1.ColumnSet = Col1;
                    Qry1.Criteria = new FilterExpression(LogicalOperator.And);
                    Qry1.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                    Qry1.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, iPin.Id));
                    EntityCollection Colec1 = service.RetrieveMultiple(Qry1);
                    if (Colec1.Entities.Count > 0)
                    {
                        hil_businessmapping iBus = Colec1.Entities[0].ToEntity<hil_businessmapping>();
                        _CstA["hil_pincode"] = new EntityReference(hil_businessmapping.EntityLogicalName, iBus.Id);
                    }
                }     
            }
            if(InvoiceDate != null)
            {
                _CstA.hil_InvoiceDate = InvoiceDate.AddDays(1);
            }
            if(InvoiceName != null)
            {
               _CstA.hil_InvoiceNo = InvoiceName;
            }
            if (DummyAc != Guid.Empty)
            {
                _CstA.msdyn_Account = new EntityReference(Account.EntityLogicalName, DummyAc);
            }
            if(Stage.hil_ProductCategoryDivision != null && Stage.hil_ProductSubCategoryMG != null)
            {
                _CstA.hil_ProductCategory = Stage.hil_ProductCategoryDivision;
                _CstA.hil_ProductSubcategory = Stage.hil_ProductSubCategoryMG;
                _CstA.hil_productsubcategorymapping = new EntityReference(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, Stage.Id);
            }
            if(Brdg.Contains("hil_productlocation"))
            {
                if(Brdg.Attributes.Contains("hil_productlocation"))
                {
                    int opLoc = Brdg.GetAttributeValue<int>("hil_productlocation");
                    if(opLoc > 0)
                    {
                        _CstA.hil_Product = new OptionSetValue(opLoc);
                    }
                }
            }
            if(Brdg.hil_NatureofCompalint != null)
            {
                _CstA["hil_purchasedfrom"] = Brdg.hil_NatureofCompalint;
            }
            _CstA.hil_CreateWarranty = false;
            _CstA["hil_source"] = new OptionSetValue(1);
            VD.CustomerAssetId = service.Create(_CstA);
            if(FncCode == 5 && VD.CustomerAssetId != Guid.Empty)
            {
                AttachNotes(service, Base64String, VD.CustomerAssetId, FileType);
                VD.StatusCode = true;
            }
            return (VD);
            //if(FncCode == 5)
            //{
            //    Guid WorkOd = GetThisWorkOrder(service, VD.CustomerAssetId);
            //    if (WorkOd != Guid.Empty)
            //    {
            //        AttachNotes(service, Base64String, WorkOd, FileType);
            //        VD.StatusCode = true;
            //    }
            //}
        }
        public static Guid GetThisWorkOrder(IOrganizationService service, Guid Asst)
        {
            Guid WorkOrder = new Guid();
            QueryByAttribute Query = new QueryByAttribute();
            Query.EntityName = msdyn_workorder.EntityLogicalName;
            Query.AddAttributeValue("msdyn_customerasset", Asst);
            ColumnSet col = new ColumnSet(true);
            Query.ColumnSet = col;
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorder Wo in Found.Entities)
                {
                    WorkOrder = Wo.Id;
                }
            }
            return (WorkOrder);
        }
        #endregion
        #region Attach Notes
        public static void AttachNotes(IOrganizationService service, string Notes, Guid Asset, int fileType)
        {
            try
            {
                if ((Notes != null) && (fileType != null))//int can't be null
                {
                    if (fileType == 0)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "image/jpeg";
                        An.FileName = "invoice.jpeg";
                        An.ObjectId = new EntityReference(msdyn_customerasset.EntityLogicalName, Asset);
                        An.ObjectTypeCode = msdyn_customerasset.EntityLogicalName;
                        service.Create(An);
                    }
                    else if (fileType == 1)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "image/png";
                        An.FileName = "invoice.png";
                        An.ObjectId = new EntityReference(msdyn_customerasset.EntityLogicalName, Asset);
                        An.ObjectTypeCode = msdyn_customerasset.EntityLogicalName;
                        service.Create(An);
                    }
                    else if (fileType == 2)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "application/pdf";
                        An.FileName = "invoice.pdf";
                        An.ObjectId = new EntityReference(msdyn_customerasset.EntityLogicalName, Asset);
                        An.ObjectTypeCode = msdyn_customerasset.EntityLogicalName;
                        service.Create(An);
                    }
                    else if (fileType == 3)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "application/doc";
                        An.FileName = "invoice.doc";
                        An.ObjectId = new EntityReference(msdyn_customerasset.EntityLogicalName, Asset);
                        An.ObjectTypeCode = msdyn_customerasset.EntityLogicalName;
                        service.Create(An);
                    }
                    else if (fileType == 4)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "application/tiff";
                        An.FileName = "invoice.tiff";
                        An.ObjectId = new EntityReference(msdyn_customerasset.EntityLogicalName, Asset);
                        An.ObjectTypeCode = msdyn_customerasset.EntityLogicalName;
                        service.Create(An);
                    }
                    else if (fileType == 5)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "application/gif";
                        An.FileName = "invoice.gif";
                        An.ObjectId = new EntityReference(msdyn_customerasset.EntityLogicalName, Asset);
                        An.ObjectTypeCode = msdyn_customerasset.EntityLogicalName;
                        service.Create(An);
                    }
                    else if (fileType == 6)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "application/bmp";
                        An.FileName = "invoice.bmp";
                        An.ObjectId = new EntityReference(msdyn_customerasset.EntityLogicalName, Asset);
                        An.ObjectTypeCode = msdyn_customerasset.EntityLogicalName;
                        service.Create(An);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error Attaching Notes" + ex.Message);
            }
        }
        #endregion
        #region Create WorkOrder For Demo
        public static Validate CreateWorkOrderForDemo(IOrganizationService service,string ContactGuID, Guid ServiceAccount, string StagingProdSBU, string CallSubType, string Desc)
        {
            Validate VD = new Validate();
            msdyn_workorder _WrkOd = new msdyn_workorder();
            hil_callsubtype iCall = (hil_callsubtype)service.Retrieve(hil_callsubtype.EntityLogicalName, new Guid(CallSubType), new ColumnSet("hil_name"));
            Guid PriceList = Helper.GetGuidbyName(PriceLevel.EntityLogicalName, "name", "Default Price List", service);
            _WrkOd.hil_CustomerRef = new EntityReference(Contact.EntityLogicalName, new Guid(ContactGuID));
            _WrkOd.msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
            _WrkOd.msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
            if(CallSubType != String.Empty)
            {
                _WrkOd.hil_CallSubType = new EntityReference(hil_callsubtype.EntityLogicalName, new Guid(CallSubType));
            }
            if (PriceList != Guid.Empty)
            {
                _WrkOd.msdyn_PriceList = new EntityReference(PriceLevel.EntityLogicalName, PriceList);
            }
            EntityReference CallType = new EntityReference(hil_calltype.EntityLogicalName);
            hil_callsubtype Call = (hil_callsubtype)service.Retrieve(hil_callsubtype.EntityLogicalName,new Guid(CallSubType), new ColumnSet(true));
            if(Call.hil_CallType != null)
            {
                CallType = Call.hil_CallType;
                _WrkOd.hil_CallType = new EntityReference(hil_calltype.EntityLogicalName, CallType.Id);
            }
            hil_stagingdivisonmaterialgroupmapping Stage = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, new Guid(StagingProdSBU), new ColumnSet("hil_productcategorydivision", "hil_productsubcategorymg"));

            if(Stage.hil_ProductCategoryDivision != null && Stage.hil_ProductSubCategoryMG != null)
            {
                _WrkOd.hil_Productcategory = Stage.hil_ProductCategoryDivision;
                _WrkOd.hil_ProductCatSubCatMapping = new EntityReference(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, new Guid(StagingProdSBU));
                _WrkOd.hil_ProductSubcategory = Stage.hil_ProductSubCategoryMG;
                Guid iNature = GetNature(service, iCall.hil_name, Stage.hil_ProductSubCategoryMG);
                if (iNature != Guid.Empty)
                    _WrkOd.hil_natureofcomplaint = new EntityReference(hil_natureofcomplaint.EntityLogicalName, iNature);
            }
            _WrkOd.hil_quantity = 1;
            _WrkOd.hil_SourceofJob = new OptionSetValue(4);
            _WrkOd.hil_CustomerComplaintDescription = Desc;
            Guid _WoID = service.Create(_WrkOd);
            if (_WoID != Guid.Empty)
            {
                VD.StatusCode = true;
                msdyn_workorder Wrk = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, _WoID, new ColumnSet("msdyn_name"));
                VD.WOUniqueID = Wrk.msdyn_name;
                VD.WOUniqueReferenceNumber = Convert.ToString(_WoID);
                VD.FunctionCode = 6;
                VD.ContactGuID = ContactGuID;
                VD.ExceptionCode = "XX";
                VD.ExceptionDesc = "XX";
            }
            else
            {
                VD.StatusCode = false;
                VD.ExceptionCode = "10";
                VD.ExceptionDesc = "CRM Error";
                VD.FunctionCode = 6;
                VD.WOUniqueID = "Not Created";
                VD.ContactGuID = ContactGuID;
            }
            return (VD);
        }
        #endregion
        #region Create Work Order For Service
        public static Validate CreateWorkOrderForService(IOrganizationService service, string ContactGuID, string NatureDesc, string CallSbTyp, string StagingProdSBU, string Desc, string Asset)
        {
            Validate VD = new Validate();
            msdyn_workorder _WrkOd = new msdyn_workorder();
            Guid PriceList = Helper.GetGuidbyName(PriceLevel.EntityLogicalName, "name", "Default Price List", service);
            Guid Nature = Helper.GetGuidbyName(hil_natureofcomplaint.EntityLogicalName, "hil_name", "Service", service);
            Guid IncidentType = Helper.GetGuidbyName(msdyn_incidenttype.EntityLogicalName, "msdyn_name", "Service", service);
            Guid ServiceAccount = Helper.GetGuidbyName(Account.EntityLogicalName, "name", "Dummy Customer", service);
            //Guid CallSubType = Helper.GetGuidbyName(hil_callsubtype.EntityLogicalName, "hil_name", "Break Down", service);
            _WrkOd.hil_CustomerRef = new EntityReference(Contact.EntityLogicalName, new Guid(ContactGuID));
            if(StagingProdSBU != null)
            {
                hil_stagingdivisonmaterialgroupmapping Stage = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, new Guid(StagingProdSBU), new ColumnSet("hil_productcategorydivision", "hil_productsubcategorymg"));
                if (Stage.hil_ProductCategoryDivision != null && Stage.hil_ProductSubCategoryMG != null)
                {
                    _WrkOd.hil_Productcategory = Stage.hil_ProductCategoryDivision;
                    _WrkOd.hil_ProductSubcategory = Stage.hil_ProductSubCategoryMG;
                    _WrkOd.hil_ProductCatSubCatMapping = new EntityReference(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, new Guid(StagingProdSBU));
                }
            }
            if(Asset != null)
            {
                _WrkOd.msdyn_CustomerAsset = new EntityReference(msdyn_customerasset.EntityLogicalName, new Guid(Asset));
                msdyn_customerasset iAsst = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, new Guid(Asset), new ColumnSet("msdyn_product", "hil_productsubcategory", "hil_productsubcategorymapping", "hil_productcategory"));
                if(iAsst.hil_ProductCategory != null)
                {
                    _WrkOd.hil_Productcategory = iAsst.hil_ProductCategory;
                }
                if(iAsst.hil_ProductSubcategory != null)
                {
                    _WrkOd.hil_ProductSubcategory = iAsst.hil_ProductSubcategory;
                }
                if(iAsst.hil_productsubcategorymapping != null)
                {
                    _WrkOd.hil_ProductCatSubCatMapping = iAsst.hil_productsubcategorymapping;
                }
            }
            if (ServiceAccount != Guid.Empty)
            {
                _WrkOd.msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
                _WrkOd.msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
            }
            _WrkOd.hil_CallSubType = new EntityReference(hil_callsubtype.EntityLogicalName, new Guid(CallSbTyp));
            if (PriceList != Guid.Empty)
            {
                _WrkOd.msdyn_PriceList = new EntityReference(PriceLevel.EntityLogicalName, PriceList);
            }
            if (Nature != Guid.Empty)
            {
                _WrkOd.hil_natureofcomplaint = new EntityReference(hil_natureofcomplaint.EntityLogicalName, Nature);
            }
            if (IncidentType != Guid.Empty)
            {
                _WrkOd.msdyn_PrimaryIncidentType = new EntityReference(msdyn_incidenttype.EntityLogicalName, IncidentType);
            }
            EntityReference CallType = new EntityReference(hil_calltype.EntityLogicalName);
            hil_callsubtype Call = (hil_callsubtype)service.Retrieve(hil_callsubtype.EntityLogicalName, new Guid(CallSbTyp), new ColumnSet(true));
            if (Call.hil_CallType != null)
            {
                CallType = Call.hil_CallType;
                _WrkOd.hil_CallType = new EntityReference(hil_calltype.EntityLogicalName, CallType.Id);
            }
            if (NatureDesc != null)
            {
                _WrkOd.msdyn_PrimaryIncidentDescription = NatureDesc;
            }
            _WrkOd.hil_quantity = 1;
            _WrkOd.hil_SourceofJob = new OptionSetValue(4);
            _WrkOd.hil_CustomerComplaintDescription = Desc;
            Guid _WoID = service.Create(_WrkOd);
            if (_WoID != Guid.Empty)
            {
                VD.StatusCode = true;
                msdyn_workorder Wrk = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, _WoID, new ColumnSet("msdyn_name"));
                VD.WOUniqueID = Wrk.msdyn_name;
                VD.WOUniqueReferenceNumber = Convert.ToString(_WoID);
                VD.FunctionCode = 7;
                VD.ContactGuID = ContactGuID;
                VD.ExceptionCode = "XX";
                VD.ExceptionDesc = "XX";
            }
            else
            {
                VD.StatusCode = false;
                VD.ExceptionCode = "10";
                VD.ExceptionDesc = "CRM Error";
                VD.FunctionCode = 7;
                VD.WOUniqueID = "Not Created";
                VD.ContactGuID = ContactGuID;
            }
            return (VD);
        }
        #endregion
        #region Contact Profile Update
        public static Validate ContactProfileUpdate(IOrganizationService service, hil_consumerappbridge Brdg)
        {
            Validate VD = new Validate();
            Contact Cont = new Contact();
            Cont.ContactId = new Guid(Brdg.hil_ContactGuId);
            if (Brdg.hil_FirstName != null)
            {
                Cont.FirstName = Brdg.hil_FirstName;
            }
            if (Brdg.hil_LastName != null)
            {
                Cont.LastName = Brdg.hil_LastName;
            }
            if (Brdg.hil_MobileNumber != null)
            {
                Cont.MobilePhone = Brdg.hil_MobileNumber;
            }
            if (Brdg.hil_AddressLine1 != null)
            {
                Cont.Address1_Line1 = Brdg.hil_AddressLine1;
            }
            if(Brdg.Attributes.Contains("hil_salutationcode"))
            {
                int Salutation = (int)Brdg["hil_salutationcode"];
                Cont.hil_Salutation = new OptionSetValue(Salutation);
            }
            if (Brdg.hil_AddressLine2 != null)
            {
                Cont.Address1_Line2 = Brdg.hil_AddressLine2;
            }
            if (Brdg.hil_AddressLine3 != null)
            {
                Cont.Address1_Line3 = Brdg.hil_AddressLine3;
            }
            if (Brdg.hil_City != null)
            {
                Cont.hil_city = new EntityReference(hil_city.EntityLogicalName, new Guid(Brdg.hil_City));
            }
            if (Brdg.hil_PinCode != null)
            {
                Cont.hil_pincode = new EntityReference(hil_pincode.EntityLogicalName, new Guid(Brdg.hil_PinCode));
            }
            if (Brdg.hil_State != null)
            {
                Cont.hil_state = new EntityReference(hil_state.EntityLogicalName, new Guid(Brdg.hil_State));
            }
            if (Brdg.hil_ShippingAddressLine1 != null)
            {
                Cont.Address2_Line1 = Brdg.hil_ShippingAddressLine1;
            }
            if (Brdg.hil_ShippingAddressLine2 != null)
            {
                Cont.Address2_Line2 = Brdg.hil_ShippingAddressLine2;
            }
            if (Brdg.hil_ShippingAddressLine3 != null)
            {
                Cont.Address2_Line3 = Brdg.hil_ShippingAddressLine3;
            }
            if (Brdg.hil_ShippingCity != null)
            {
                Cont.hil_ShippingCity = new EntityReference(hil_city.EntityLogicalName, new Guid(Brdg.hil_ShippingCity));
            }
            if (Brdg.hil_ShippingPincode != null)
            {
                Cont.hil_ShippingPinCode = new EntityReference(hil_pincode.EntityLogicalName, new Guid(Brdg.hil_ShippingPincode));
            }
            if (Brdg.hil_ShippingState != null)
            {
                Cont.hil_ShippingState = new EntityReference(hil_state.EntityLogicalName, new Guid(Brdg.hil_ShippingState));
            }
            service.Update(Cont);
            VD.StatusCode = true;
            VD.ExceptionCode = "XX";
            VD.ExceptionDesc = "XX";
            VD.FunctionCode = 8;
            VD.ContactGuID = Convert.ToString(Cont.Id);
            return (VD);
        }
        #endregion
        #region Get IV and Encryption Key From Integration Configuration
        public static string DecryptThisPassword(IOrganizationService service, string Pwd)
        {
            string _Decrypted = string.Empty;
            string keyString = string.Empty;
            string _InVector = string.Empty;
            byte[] Key = new byte[] { 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20 };
            byte[] IntiVector = new byte[] { 0x20, 0x20, 0x20, 0x20, 0x20, 0x20, 0x20 };
            QueryByAttribute Qry = new QueryByAttribute(hil_integrationconfiguration.EntityLogicalName);
            ColumnSet Col = new ColumnSet("hil_origin", "hil_signature");
            Qry.AddAttributeValue("hil_name", "Password Encryption");
            EntityCollection Found = new EntityCollection();
            Found = service.RetrieveMultiple(Qry);
            if (Found.Entities.Count == 1)
            {
                foreach (hil_integrationconfiguration _IntConf in Found.Entities)
                {
                    keyString = _IntConf.hil_Origin;
                    _InVector = _IntConf.hil_Signature;
                    Key = Encoding.ASCII.GetBytes(keyString);
                    IntiVector = Encoding.ASCII.GetBytes(_InVector);
                }
            }
            _Decrypted = Decrypt(Pwd, Key, IntiVector);
            return (_Decrypted);
        }
        #endregion
        #region Decrypt Password AES 256
        private static string Decrypt(string cipherText, byte[] key, byte[] iv)
        {
            // Instantiate a new Aes object to perform string symmetric encryption
            Aes encryptor = Aes.Create();

            encryptor.Mode = CipherMode.CBC;

            // Set key and IV
            encryptor.Key = key;
            encryptor.IV = iv;

            // Instantiate a new MemoryStream object to contain the encrypted bytes
            MemoryStream memoryStream = new MemoryStream();

            // Instantiate a new encryptor from our Aes object
            ICryptoTransform aesDecryptor = encryptor.CreateDecryptor();

            // Instantiate a new CryptoStream object to process the data and write it to the 
            // memory stream
            CryptoStream cryptoStream = new CryptoStream(memoryStream, aesDecryptor, CryptoStreamMode.Write);

            // Will contain decrypted plaintext
            string plainText = String.Empty;

            try
            {
                // Convert the ciphertext string into a byte array
                byte[] cipherBytes = Convert.FromBase64String(cipherText);

                // Decrypt the input ciphertext string
                cryptoStream.Write(cipherBytes, 0, cipherBytes.Length);

                // Complete the decryption process
                cryptoStream.FlushFinalBlock();

                // Convert the decrypted data from a MemoryStream to a byte array
                byte[] plainBytes = memoryStream.ToArray();

                // Convert the decrypted byte array to string
                plainText = Encoding.ASCII.GetString(plainBytes, 0, plainBytes.Length);
            }
            finally
            {
                // Close both the MemoryStream and the CryptoStream
                memoryStream.Close();
                cryptoStream.Close();
            }
            return plainText;
        }
        #endregion
        #region Update Bridge Table
    private static void UpdateBridgeRecord(IOrganizationService service, Validate VD, hil_consumerappbridge Brdg, ITracingService Tracing)
        {
            switch (VD.FunctionCode)
            {
                case 1://CREATE CUSTOMER
                    if (VD.StatusCode == true)
                    {
                        Brdg.hil_ContactGuId = VD.ContactGuID;
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_ContactRef = new EntityReference(Contact.EntityLogicalName, new Guid(VD.ContactGuID));
                        Brdg.hil_ExceptionCode = "XX";
                        Brdg.hil_ExceptionDetail = "XX";
                    }
                    else
                    {
                        Brdg.hil_ContactGuId = "NOT FOUND";
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_ExceptionCode = VD.ExceptionCode;
                        Brdg.hil_ExceptionDetail = VD.ExceptionDesc;
                    }
                    break;
                case 2://VALIDATE CUSTOMER
                    if (VD.StatusCode == true)
                    {
                        Brdg.hil_ContactGuId = VD.ContactGuID;
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_IfValidated = VD.IfValidated;
                        if(VD.IfValidated == true)
                        {
                            Brdg.hil_ContactRef = new EntityReference(Contact.EntityLogicalName, new Guid(VD.ContactGuID));
                        }
                        Brdg.hil_FirstName = VD.FirstName;
                        Brdg.hil_LastName = VD.LastName;
                        Brdg.hil_MobileNumber = VD.MobileNumber;
                        Brdg.hil_PinCode = VD.PinCode;
                        Brdg.hil_ExceptionCode = "XX";
                        Brdg.hil_ExceptionDetail = "XX";
                    }
                    else
                    {
                        Brdg.hil_IfValidated = false;
                        Brdg.hil_MobileNumber = "NOT FOUND";
                        Brdg.hil_ContactGuId = "NOT FOUND";
                        Brdg.hil_FirstName = "NOT FOUND";
                        Brdg.hil_LastName = "NOT FOUND";
                        Brdg.hil_PinCode = "NOT FOUND";
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_ExceptionCode = VD.ExceptionCode;
                        Brdg.hil_ExceptionDetail = VD.ExceptionDesc;
                    }
                    break;
                case 3://FORGOT PASSWORD
                    if (VD.StatusCode == true)
                    {
                        Brdg.hil_ContactGuId = VD.ContactGuID;
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_OTP = VD.OTP;
                        Brdg.hil_ContactRef = new EntityReference(Contact.EntityLogicalName, new Guid(VD.ContactGuID));
                        Brdg.hil_ExceptionCode = "XX";
                        Brdg.hil_ExceptionDetail = "XX";
                    }
                    else
                    {
                        Brdg.hil_OTP = "XX";
                        Brdg.hil_ContactGuId = "NOT FOUND";
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_ExceptionCode = VD.ExceptionCode;
                        Brdg.hil_ExceptionDetail = VD.ExceptionDesc;
                    }
                    break;
                case 4://CHANGE PASSWORD
                    if (VD.StatusCode == true)
                    {
                        Brdg.hil_ContactGuId = VD.ContactGuID;
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_ContactRef = new EntityReference(Contact.EntityLogicalName, new Guid(VD.ContactGuID));
                        Brdg.hil_ExceptionCode = "XX";
                        Brdg.hil_ExceptionDetail = "XX";
                    }
                    else
                    {
                        Brdg.hil_ContactGuId = "Not Found";
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_ExceptionCode = VD.ExceptionCode;
                        Brdg.hil_ExceptionDetail = VD.ExceptionDesc;
                    }
                    break;
                case 5://PRODUCT REGISTRATION
                    if (VD.StatusCode == true)
                    {
                        Tracing.Trace(Convert.ToString(DateTime.Now));
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_ExceptionCode = "XX";
                        Brdg.hil_ExceptionDetail = "XX";
                    }
                    else
                    {
                        Tracing.Trace(Convert.ToString(DateTime.Now));
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_ExceptionCode = VD.ExceptionCode;
                        Brdg.hil_ExceptionDetail = VD.ExceptionDesc;
                    }
                    break;
                case 6://CREATE WORK ORDER DEMO
                    if(VD.StatusCode == true)
                    {
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_WorkOrderGuId = VD.WOUniqueID;
                        Brdg.hil_ContactGuId = VD.ContactGuID;
                        Brdg.hil_ExceptionCode = "XX";
                        Brdg.hil_ExceptionDetail = "XX";
                    }
                    else
                    {
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_WorkOrderGuId = "Not Found";
                        Brdg.hil_ContactGuId = VD.ContactGuID;
                        Brdg.hil_ExceptionCode = VD.ExceptionCode;
                        Brdg.hil_ExceptionDetail = VD.ExceptionDesc;
                    }
                    break;
                case 7://CREATE WORK ORDER SERVICE
                    if (VD.StatusCode == true)
                    {
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_WorkOrderGuId = VD.WOUniqueID;
                        Brdg.hil_ContactGuId = VD.ContactGuID;
                        Brdg.hil_ExceptionCode = "XX";
                        Brdg.hil_ExceptionDetail = "XX";
                    }
                    else
                    {
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_WorkOrderGuId = "Not Found";
                        Brdg.hil_ContactGuId = VD.ContactGuID;
                        Brdg.hil_ExceptionCode = VD.ExceptionCode;
                        Brdg.hil_ExceptionDetail = VD.ExceptionDesc;
                    }
                    break;
                case 8://PROFILE UPDATE
                    if (VD.StatusCode == true)
                    {
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_ExceptionCode = "XX";
                        Brdg.hil_ExceptionDetail = "XX";
                        Brdg.hil_ContactGuId = VD.ContactGuID;
                    }
                    else
                    {
                        Brdg.hil_StatusCode = VD.StatusCode;
                        Brdg.hil_ExceptionCode = VD.ExceptionCode;
                        Brdg.hil_ExceptionDetail = VD.ExceptionDesc;
                        Brdg.hil_ContactGuId = "Not Found";
                    }
                    break;
            }
            //service.Update(Brdg);
        }
        #endregion
        #region Check If Contact Exists
        public static Guid GetThisCustomerIfExists(IOrganizationService service, string Mobile, string Email)
        {
            Guid ContactId = Guid.Empty;
            QueryExpression Query = new QueryExpression(Contact.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.Or);
            Query.Criteria.AddCondition(new ConditionExpression("mobilephone", ConditionOperator.Equal, Mobile));
            Query.Criteria.AddCondition(new ConditionExpression("emailaddress1", ConditionOperator.Equal, Email));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count > 0)
            {
                foreach(Contact cnt in Found.Entities)
                {
                    ContactId = cnt.Id;
                }
            }
            return ContactId;
        }
        #endregion
        #region Get PinCode
        public static string GetPinCode(IOrganizationService service, Guid ContId)
        {
            string PinCode = " ";
            QueryByAttribute Query = new QueryByAttribute(hil_address.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_pincode");
            Query.AddAttributeValue("hil_customer", ContId);
            Query.AddAttributeValue("hil_addresstype", 1);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count > 0)
            {
                hil_address Add = (hil_address)Found.Entities[0];
                PinCode = Add.hil_PinCode.Name;
            }
            return PinCode;
        }
        #endregion
        #region Get Nature
        public static Guid GetNature(IOrganizationService service, string Call, EntityReference ProductSCat)
        {
            Guid Nature = new Guid();
            QueryExpression Query = new QueryExpression(hil_natureofcomplaint.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, Call));
            Query.Criteria.AddCondition(new ConditionExpression("hil_relatedproduct", ConditionOperator.Equal, ProductSCat.Id));
            Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count > 0)
            {
                Nature = Found.Entities[0].Id;
            }
            return Nature;
        }
        #endregion
        #region Get Business Mapping
        public static hil_businessmapping GetBusinessMapping(IOrganizationService service, Guid PinCodeId)
        {
            hil_businessmapping iMap = new hil_businessmapping();
            QueryExpression Query = new QueryExpression(hil_businessmapping.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, PinCodeId));
            Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count > 0)
            {
                iMap = Found.Entities[0].ToEntity<hil_businessmapping>();
            }
            return iMap;
        }
        #endregion
    }
}