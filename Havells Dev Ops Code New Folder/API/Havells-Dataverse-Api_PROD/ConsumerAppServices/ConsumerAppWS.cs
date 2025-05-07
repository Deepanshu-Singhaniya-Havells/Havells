using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Net;
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;
using ConsumerApp.BusinessLayer;
using System.Web;
//using System.Web.Http;

namespace ConsumerAppServices
{
    //NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IService1" in both code and config file together.
    [ServiceContract]
    public interface IConsumerAppWS
    {
        [OperationContract]
        string GetData(int value);

        [OperationContract]
        [WebInvoke(UriTemplate = "/GetJobs",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<JobOutput> GetJobs(Job InputJob);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppCreateAddress",
            Method = "POST", RequestFormat = WebMessageFormat.Json,
            ResponseFormat = WebMessageFormat.Json)]
        ConsumerAddress ConsumerAppAddressCreate(ConsumerAddress Addr);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppValidatedOTP",
            Method = "POST", RequestFormat = WebMessageFormat.Json,
            ResponseFormat = WebMessageFormat.Json)]
        CreateContactFromBridge ConsumerAppValidatedOTP(CreateContactFromBridge CrBridge);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppForgotPassword",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        ForgotPassword ConsumerAppForgotPassword(ForgotPassword forgotPassword);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppForgotPasswordOTP",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        ForgotPasswordCheck1 ConsumerAppForgotPasswordCheck1(ForgotPasswordCheck1 forgotPassword);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppRegisterCustomer",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        RegisterCustomer ConsumerAppRegisterCustomer(RegisterCustomer Register);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppAuthenticateCustomer",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        AuthenticateCustomer ConsumerAppAuthenticateCustomer(AuthenticateCustomer Authenticate);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppAuthenticateCustomerDE",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        AuthenticateCustomer ConsumerAppAuthenticateCustomerDE(AuthenticateCustomer Authenticate);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppCityMaster",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<CityMaster> ConsumerAppCityMaster(CityMaster City);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppPinCodeMaster",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<PinCodeMaster> ConsumerAppPincodeMaster(PinCodeMaster Pin);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppExistingServices",
                   Method = "POST", ResponseFormat = WebMessageFormat.Json,
                   RequestFormat = WebMessageFormat.Json)]
        List<ExistingServices> ConsumerAppExistingServices(ExistingServices Serv);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppPinCodeFilter",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<PinCodeFilter> ConsumerAppPincodeFilter(PinCodeFilter Pin);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppGeoPinCode",
                   Method = "GET", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<PinCodeMasterData> ConsumerAppPinCodeMaster();
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppStateMaster",
                   Method = "GET", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<StateMaster> ConsumerAppStateMaster();
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppCityMasterUnfiltered",
                   Method = "GET", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<CityUnFiltered> ConsumerAppCityUnfilter();
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppCallSubTypeMaster",
                   Method = "GET", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<CallSubtypeMaster> ConsumerAppCallSubTypeMaster();
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppCallSubTypeService",
                   Method = "GET", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<CallSubtypeMasterForService> CallSubtypeMasterService();
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppChangePassword",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        ChangePassword ConsumerAppChangePassword(ChangePassword details);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppProfileUpdate",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        ProfileUpdate ConsumerAppProfileUpdate(ProfileUpdate details);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppProductRegistration",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        ProductRegistration ConsumerAppProductRegistration(ProductRegistration ProductRegister);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppCustomerWishList",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<CustomerWishList> ConsumerAppCustomerWishList(CustomerWishList customerWishList);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppCreateWorkOrder",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        CreateWorkOrder ConsumerAppCreateWorkOrder(CreateWorkOrder bridge);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppCreateWorkOrderForDemo",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        CreateWorkOrderForDemo ConsumerAppCreateWorkOrderForDemo(CreateWorkOrderForDemo bridge);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppProductCategorySubCategory",
                   Method = "GET", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<ProductCategorySubCategoryMaster> ConsumerAppProductCategorySubCategory();
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppAlertsAndNotifications",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<AlertsAndNotifications> ConsumerAppAlertsAndNotifications(AlertsAndNotifications bridge);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppAlertsAndNotificationsPerDay",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<AlertsAndNotificationsPerDay> ConsumerAppAlertsAndNotificationsPerDay(AlertsAndNotificationsPerDay bridge);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppExisitingPurchases",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<ExisitingPurchases> ConsumerAppExistingPurchases(ExisitingPurchases bridge);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppGeoLocator",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<GeoLocator> ConsumerAppGeoLocator(GeoLocator bridge);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppCustomerProfileDetails",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        ProfileView ConsumerAppCustomerProfileDetails(ProfileView bridge);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppGetCustomerGuid",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        GetCustomerGuid ConsumerAppGetCustomerGuid(GetCustomerGuid bridge);
        [OperationContract]
        [WebInvoke(UriTemplate = "/AVAYAGetCustomerFullName",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        GetCustomerName AVAYAGetCustomerFullName(GetCustomerName Cust);
        [OperationContract]
        [WebInvoke(UriTemplate = "/AVAYAGetLastOpenCase",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        GetLastOpenCase AVAYAGetLastOpenCase(GetLastOpenCase Cust);
        [OperationContract]
        [WebInvoke(UriTemplate = "/AVAYAEscallateWrkOrder",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        EscalateWorkOrder AVAYAEscallateWrkOrder(EscalateWorkOrder Cust);
        [OperationContract]
        [WebInvoke(UriTemplate = "/AVAYACreateWO",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        AVAYACreateWo AVAYACreateWO(AVAYACreateWo Cust);
        [OperationContract]
        [WebInvoke(UriTemplate = "/PushCustomerECommerce",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        ReturnInfo ECommercePushCust(PUSH_Customer Cust);

        [OperationContract]
        [WebGet(UriTemplate = "/ConsumerAppInSMS?FROM={FROM}&TO={TO}&MESSAGE={MESSAGE}&JobType={JobType}", BodyStyle = WebMessageBodyStyle.Bare,
                 RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)
             ]
        IN_SMS InComingSMSCreate(string FROM, string TO, string MESSAGE, int JobType);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppPushLead",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        PUSH_LEAD CreateLeadAvaya(PUSH_LEAD iLead);
        [OperationContract]
        [WebInvoke(UriTemplate = "/GetJobDetails",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        JobDetails FindJobDetails(JobDetails iDetails);
        [OperationContract]
        [WebInvoke(UriTemplate = "/GetProductSubCategory",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<ProductSubCategoryMaster> GetProductSubCategory(ProductSubCategoryMaster iDetails);
        [OperationContract]
        [WebInvoke(UriTemplate = "/GetProductMaster",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<ProductCodeMaster> GetProductMaster(ProductCodeMaster iDetails);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ShootSMS",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        OUTGOING_SMS SendSMS(OUTGOING_SMS iDetails);
        [OperationContract]
        [WebInvoke(UriTemplate = "/AuthenticateOTP",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        AuthenticateCustomerOTP iAuthenticate(AuthenticateCustomerOTP iDetails);
        [OperationContract]
        [WebInvoke(UriTemplate = "/PostFeedback",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        Feedback iFeedback(Feedback iDetails);
        [OperationContract]
        [WebInvoke(UriTemplate = "/GetModelDetails",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        FindMaterialByName GetMaterialDetailsbyName(FindMaterialByName iDetails);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppValidateSerialNumber",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        ValidateSerialNumber iValidateSerialNumFromSAP(ValidateSerialNumber iValidate);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppValidateSerialNumProductDelink",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        ValidateSerialNumber ValidateSerialNumProductDelink(ValidateSerialNumber iValidate);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppValidateUser",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        IsExistingUser UserValidation(IsExistingUser iCheck);
        [OperationContract]
        [WebInvoke(UriTemplate = "/FeatureInMaking",
                   Method = "GET", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        bool FeatureInMaking();
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppGetAddressList",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<OutputAddressClass> GetAllUserAddress(GetUserAddresses iConsumer);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppViewAddress",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        AddressFields ViewThisAddress(ViewAddress iViewAddress);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppAddNewAddress",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        OutputAddAddress AddConsumerAddress(AddAddress iAdd);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppAddUpdateAddress",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        OutputUpdateAddress UpdateConsumerAddress(UpdateAddress iUpdate);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppCreateServiceTicketNew",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        OutputCreateServiceTicket_Breakdown CreateServiceTicketNew(CreateServiceTicket_Breakdown iCreateTicket);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppDeleteAddress",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        OutputAddressDelete DeleteCustomerAddress(AddressDelete iDeleteAdd);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppValidateSerialNumberQA",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        ValidateSerialNumberQAcs iValidateSerialNumFromSAPQA(ValidateSerialNumberQAcs iValidate);
        [OperationContract]
        [WebInvoke(UriTemplate = "/ConsumerAppGetDistricts",
            Method = "POST", RequestFormat = WebMessageFormat.Json,
            ResponseFormat = WebMessageFormat.Json)]
        List<OutputDistrictMaster> GetDistrictMaster(DistrictMaster iMaster);
        [OperationContract]
        [WebInvoke(UriTemplate = "/CreateWorkOrderQA",
            Method = "POST", RequestFormat = WebMessageFormat.Json,
            ResponseFormat = WebMessageFormat.Json)]
        CreateWorkOrderQA CreateJobsinQA(CreateWorkOrderQA iMaster);

        [OperationContract]
        [WebInvoke(UriTemplate = "/Authenticate",
            Method = "POST", RequestFormat = WebMessageFormat.Json,
            ResponseFormat = WebMessageFormat.Json)]
        AuthenticationServiceExternal Authenticate(AuthenticationServiceExternal oSFA);

        [OperationContract]
        [WebInvoke(UriTemplate = "/CreateEnquiry",
           Method = "POST", RequestFormat = WebMessageFormat.Json,
           ResponseFormat = WebMessageFormat.Json)]
        ReturnInfoVar CreateEnquiryEntry(CreateEnquiry enquiry);

        [OperationContract]
        [WebInvoke(UriTemplate = "/RegisteredProducts",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        List<RegisteredProducts> GetRegisteredProducts(RegisteredProducts registeredProducts);

        [OperationContract]
        [WebInvoke(UriTemplate = "/NatureOfComplaint",
           Method = "POST", RequestFormat = WebMessageFormat.Json,
           ResponseFormat = WebMessageFormat.Json)]
        List<NatureOfComplaint> GetNatureOfComplaints(NatureOfComplaint natureOfComplaint);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTNatureOfComplaintByProdSubcategory",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<IoTNatureofComplaint> GetIoTNatureOfComplaintsByProdSubCatg(IoTNatureofComplaint natureOfComplaint);

        [OperationContract]
        [WebInvoke(UriTemplate = "/AddressBook",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<AddressBook> GetAddress(AddressBook addressBook);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ServiceCall",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ServiceCall CreateServiceCall(ServiceCall serviceCallData);

        [OperationContract]
        [WebInvoke(UriTemplate = "/YoutubeApi",
          Method = "GET", RequestFormat = WebMessageFormat.Json,
          ResponseFormat = WebMessageFormat.Json)]
        YoutubeApi iYoutube();

        [OperationContract]
        [WebInvoke(UriTemplate = "/TechnicianAuthentication",
           Method = "POST", RequestFormat = WebMessageFormat.Json,
           ResponseFormat = WebMessageFormat.Json)]
        TechnicianInfo GetTechnicianInfo(TechnicianInfo care360ID);

        #region WhatsApp/Havells IoT Platform Integration
        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTRegisterConsumer",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ReturnResult RegisterConsumer(IoT_RegisterConsumer consumer);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTGetCustomerProfile",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        IoTCustomerProfile GetIoTCustomerProfile(IoTCustomerProfile customerProfileData);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTUpdateCustomerProfile",
                   Method = "POST", RequestFormat = WebMessageFormat.Json,
                   ResponseFormat = WebMessageFormat.Json)]
        IoTCustomerProfile UpdateIoTCustomerProfile(IoTCustomerProfile customerProfileData);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTGetAddressBook",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<IoTAddressBookResult> GetIoTAddressBook(IoTAddressBook address);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTCreateAddress",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        IoTAddressBookResult IoTCreateAddress(IoTAddressBookResult address);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTValidatePINCode",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        IoTPINCodes IoTValidatePINCode(IoTPINCodes address);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTGetAreasByPinCode",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<IoTAreas> IoTGetAreasByPinCode(IoTPINCodes address);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTUpdateAddress",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        IoTAddressBookResult IoTUpdateAddress(IoTAddressBookResult address);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ECommercePushBulkCustomerData",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        IoTAddressBookResult ECommerceBulkCustomerData(IoTAddressBookResult addressData);

        [OperationContract]
        [WebInvoke(UriTemplate = "/RegisterProductFromWhatsapp",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        IoTRegisteredProductsResult RegisterProductFromWhatsapp(IoTRegisteredProducts productData);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTRegisterProduct",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        IoTRegisteredProductsResult IoTRegisterProduct(IoTRegisteredProducts productData);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTGetProductDetails",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        IoTValidateSerialNumber GetProductDetails(IoTValidateSerialNumber _reqParam);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTValidateSerialNumber",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        IoTValidateSerialNumber ValidateAssetSerialNumber(IoTValidateSerialNumber _reqParam);

        [OperationContract]
        [WebInvoke(UriTemplate = "/AttachNotes",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        Attachment AttachNotes(Attachment attachmentData);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTCreateServiceCall",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        IoTServiceCallRegistration IoTCreateServiceCall(IoTServiceCallRegistration serviceCalldata);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTCancelServiceCall",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        CancelJobResponse CancelServiceJob(CancelJobRequest reqParam);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTCreateServiceCallWhatsapp",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        IoTServiceCallRegistration IoTCreateServiceCallWhatsapp(IoTServiceCallRegistration serviceCalldata);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTRegisteredProducts",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<IoTRegisteredProducts> GetIoTRegisteredProducts(IoTRegisteredProducts registeredProducts);

        [OperationContract]
        [WebInvoke(UriTemplate = "/AllNatureOfComplaints",
        Method = "GET", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<NatureOfComplaint> GetAllNOCs();

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTSalutationEnum",
        Method = "GET", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<HashTableDTO> GetSalutationEnum();

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTGetServiceCalls",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<IoTServiceCallResult> GetIoTServiceCalls(IotServiceCall job);

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTInstallationLocationEnum",
        Method = "GET", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<HashTableDTO> GetInstallationLocationEnum();

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTProductHierarchy",
        Method = "GET", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<ProductHierarchyDTO> GetProductHierarchy();

        [OperationContract]
        [WebInvoke(UriTemplate = "/IoTProducts",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<ProductDTO> GetProducts(Guid productSubCategoryId);
        #endregion

        [OperationContract]
        [WebInvoke(UriTemplate = "/ValidateSerialNumberThirdParty",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ValidateSerialNumber ValidateSerialNumberThirdParty(ValidateSerialNumber iValidate);

        #region SFA Module APIs
        [OperationContract]
        [WebInvoke(UriTemplate = "/SFAValidateCustomer",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        SFA_ValidateCustomer ValidateCustomer(SFA_ValidateCustomer customerData);

        [OperationContract]
        [WebInvoke(UriTemplate = "/SFACreateServiceCall",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        SFA_ServiceCallResult SFA_CreateServiceCall(SFA_ServiceCall serviceCallData);

        [OperationContract]
        [WebInvoke(UriTemplate = "/SFAGetDivisionCallTypeSetup",
        Method = "GET", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<SFA_DivisionCallType> GetDivisionCallTypeSetup();
        #endregion

        #region D365 Audit Log Viewer
        [OperationContract]
        [WebInvoke(UriTemplate = "/GetAttributeMetadata",
           Method = "POST", RequestFormat = WebMessageFormat.Json,
           ResponseFormat = WebMessageFormat.Json)]
        List<AttributesMetadata> GetAttributeMetadata(string _entityLogicalName);
        #endregion

        #region D365 Audit Log 
        [OperationContract]
        [WebInvoke(UriTemplate = "/AttributeMetadataInfo",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<AttributeMetadataInfo> AttributeMetadata(EntityMetadataInfo entity);

        [OperationContract]
        [WebInvoke(UriTemplate = "/GetD365AuditLog",
                Method = "POST", RequestFormat = WebMessageFormat.Json,
                ResponseFormat = WebMessageFormat.Json)]
        List<D365AuditLogResult> GetD365AuditLogData(D365AuditLog _requestData);
        #endregion

        #region Tender Module
        [OperationContract]
        [WebInvoke(UriTemplate = "/GetTenderDocumentTypes",
        Method = "GET", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<DocumentTypes> GetDocumentTypes();
        #endregion

        #region Claim Extension on Tech Mobile
        [OperationContract]
        [WebInvoke(UriTemplate = "/GetWOSchemeCodes",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<WOSchemesResult> GetSchemeCodes(ClaimExtToMobile _reqParam);

        [OperationContract]
        [WebInvoke(UriTemplate = "/PatchSchemeCode",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        WOSchemesResult PatchSchemeCode(ClaimExtToMobile _reqParam);

        #endregion

        #region MRN Module APIs
        [OperationContract]
        [WebInvoke(UriTemplate = "/GetChannelPartner",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<CustomerInfo> GetCustomerDetails(string CustomerSearchString);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ValidateSerialNumberMRN",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ValidateSerialNumberMRN GetResponseFromSAP(ValidateSerialNumberMRN iValidate);

        [OperationContract]
        [WebInvoke(UriTemplate = "/DSVAddToViewList",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        MRNHeader AddtoViewList(MRNEntry _productInfo);

        [OperationContract]
        [WebInvoke(UriTemplate = "/DSVSubmitViewList",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        MRNHeader SubmitViewList(MRNHeader _viewList);

        [OperationContract]
        [WebInvoke(UriTemplate = "/DSVGetViewList",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        MRNSummary GetViewList(ViewListSearch _searchCondition);

        [OperationContract]
        [WebInvoke(UriTemplate = "/DSVGetDefectiveStockNote",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        MRNSummary GetDefectiveStockNoteForSAP(ViewListSearch _searchCondition);

        [OperationContract]
        [WebInvoke(UriTemplate = "/DSVGetPDICalls",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<PDOICalls> GetPDICalls(MRNEntry job);
        #endregion

        #region AMC Billing
        [OperationContract]
        [WebInvoke(UriTemplate = "/ValidateAMCReceiptAmount",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        AMCBilling ValidateAMCReceiptAmount(AMCBilling _reqData);
        #endregion

        #region Remote Pay
        [OperationContract]
        [WebInvoke(UriTemplate = "/SendRemotePaySMS",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        SendPaymentUrlResponse SendRemotePaySMS(SendURLD365Request reqParm);

        [OperationContract]
        [WebInvoke(UriTemplate = "/GetRemotePayStatus",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        PaymentStatusD365Response getRemotePayStatus(String jobId);

        #endregion

        // TODO: Add your service operations here
        //WOSchemesResult PatchSchemeCode(ClaimExtToMobile _reqParam)

        #region Home Advisory
        [OperationContract]
        [WebInvoke(UriTemplate = "/HVCreateAppointment",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        Response CreateAppointment(CRMRequest req);

        [OperationContract]
        [WebInvoke(UriTemplate = "/HVCancleAppointment",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        Response CancleAppointment(CancelAppointmentRequest req);

        [OperationContract]
        [WebInvoke(UriTemplate = "/HVCreateAdvisory",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        HopmeAdvisoryResult CreateAdvisory(HomeAdvisoryRequest reqParm);

        [OperationContract]
        [WebInvoke(UriTemplate = "/HVGetAdvisory",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        GetEnquiry GetEnquery(Enquiry req);

        [OperationContract]
        [WebInvoke(UriTemplate = "/GetAdvisoryEnquiryStatus",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        GetEnquiry GetEnqueryStatus(EnquiryStatus req);

        [OperationContract]
        [WebInvoke(UriTemplate = "/GetSalesEnquiryStatus",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<EnquiryDetails> GetSalesEnquiry(EnquiryStatus _enquiryStatus);

        [OperationContract]
        [WebInvoke(UriTemplate = "/HVRescheduleAppointment",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        Response RescheduleAppointment(ReschduleAppointment req);

        [OperationContract]
        [WebInvoke(UriTemplate = "/AMUploadAttachment",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ResponseUpload UploadAttachment(UploadAttachment reqParm);

        [OperationContract]
        [WebInvoke(UriTemplate = "/GetUserTimeSlots",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        GetUserTimeSlotsRoot GetUserTimeSlots(GetUserTimeSlotsRequest requestParm);

        [OperationContract]
        [WebInvoke(UriTemplate = "/GetEnquiryType",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<EnquiryType> GetEnquiryType(EnquiryType _enquiryCategory);

        [OperationContract]
        [WebInvoke(UriTemplate = "/GetEnquiryProductType",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<EnquiryProductType> GetEnquiryProductType(EnquiryProductType _enquiryCategory);

        #endregion

        #region Experience Store
        [OperationContract]
        [WebInvoke(UriTemplate = "/ESRegisterCustomer",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ExperinceZonePayLoad ESRegisterConsumer(ExperinceZonePayLoad consumer);
        #endregion

        #region Dealer Portal
        [OperationContract]
        [WebInvoke(UriTemplate = "/CreateServiceCallBase",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        IoTServiceCallRegistration IoTCreateServiceCallDealerPortal(IoTServiceCallRegistration serviceCalldata);
        #endregion

        #region Consumer NPS
        [OperationContract]
        [WebInvoke(UriTemplate = "/SendSMSEmail",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ResponseDTO SendCommunication(RequestDTO _data);
        #endregion

        #region Voice Bot APIs
        [OperationContract]
        [WebInvoke(UriTemplate = "/GetEscalations",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        Escalations GetEscalations(Escalations JobData);

        [OperationContract]
        [WebInvoke(UriTemplate = "/CreateEscalations",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        Escalations InsertEscalations(EscalationsReqRes reqParams);

        [OperationContract]
        [WebInvoke(UriTemplate = "/CreateCallbackRequest",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        CallbackRequest InsertCallbackRequest(CallbackRequest reqParams);

        [OperationContract]
        [WebInvoke(UriTemplate = "/GetProductDivision",
        Method = "GET", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ProductDivision GetProductDivisions();
        [OperationContract]
        [WebInvoke(UriTemplate = "/PostBotTranscript",
           Method = "POST", RequestFormat = WebMessageFormat.Json,
           ResponseFormat = WebMessageFormat.Json)]
        BotSpeechTransScript InsertBotSpeechTransScript(BotSpeechTransScript reqParams);

        [OperationContract]
        [WebInvoke(UriTemplate = "/CreateServiceCallVoiceBot",
           Method = "POST", RequestFormat = WebMessageFormat.Json,
           ResponseFormat = WebMessageFormat.Json)]
        IoTServiceCallRegistration IoTCreateServiceVoiceBot(IoTServiceCallRegistration serviceCalldata);
        #endregion
        [OperationContract]
        [WebInvoke(UriTemplate = "/KKGCodeVerification",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ValidateKKGCodeResult KKGCodeVerification(ValidateKKGCodeInput validateKKGCodeInput);

        [OperationContract]
        [WebInvoke(UriTemplate = "/GetOutstandingAMCs",
           Method = "POST", RequestFormat = WebMessageFormat.Json,
           ResponseFormat = WebMessageFormat.Json)]
        ServiceResponseData GetOutstandingAMCs(ReqestData reqestData);

        [OperationContract]
        [WebInvoke(UriTemplate = "/AuthServiceForAMC",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        AuthenticateConsumer AuthenticateConsumerAMC(AuthenticateConsumer requestParam);

        [OperationContract]
        [WebInvoke(UriTemplate = "/ValidateAMCSession",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ValidateSessionResponse ValidateSessionDetails(ValidateSessionRequest requestParam);

        #region APIs to retreive D365 Archived Data
        [OperationContract]
        [WebInvoke(UriTemplate = "/GetArchivedJobData",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ArchivedJobsResponse GetArchivedJobData(ArchivedJobRequest inpReq);

        [OperationContract]
        [WebInvoke(UriTemplate = "/GetArchivedData",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ResponseData GetArchivedData(RequestData inpReq);

        #endregion

        #region Airtel
        [OperationContract]
        [WebInvoke(UriTemplate = "/AIQGetOpenJobs",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ResposeDataCallMasking GetCustomerOpenJobs(RequestDataCallMasking _requestData);

        [OperationContract]
        [WebInvoke(UriTemplate = "/AIQPushCDR",
                Method = "POST", RequestFormat = WebMessageFormat.Json,
                ResponseFormat = WebMessageFormat.Json)]
        CDR_Response PushCDR(CDR_Request request);
        #endregion

        #region Consumable's History
        [OperationContract]
        [WebInvoke(UriTemplate = "/GetConsumablesHistory",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ConsumablesData RequestData(Request req);
        #endregion

        #region IOT Address Management APIs
        [OperationContract]
        [WebInvoke(UriTemplate = "V1/IoTGetAddressBook",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        List<IoTAddressBookResultV1> GetIoTAddressBookV1(IoTAddressBookV1 address);

        [OperationContract]
        [WebInvoke(UriTemplate = "V1/IoTCreateAddress",
           Method = "POST", RequestFormat = WebMessageFormat.Json,
           ResponseFormat = WebMessageFormat.Json)]
        IoTAddressBookResultV1 IoTCreateAddressV1(IoTAddressBookResultV1 address);

        [OperationContract]
        [WebInvoke(UriTemplate = "V1/IoTUpdateAddress",
           Method = "POST", RequestFormat = WebMessageFormat.Json,
           ResponseFormat = WebMessageFormat.Json)]
        IoTAddressBookResultV1 IoTUpdateAddressV1(IoTAddressBookResultV1 address);

        [OperationContract]
        [WebInvoke(UriTemplate = "V1/IoTDeleteAddress",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        IoTAddressBookResultV1 IoTDeleteAddressV1(IoTAddressBookResultV1 address);
        #endregion

        [OperationContract]
        [WebInvoke(UriTemplate = "/ValidateProductInstallation",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ValidateProductInstallation ValidateProductInstallation(ValidateProductInstallation _data);

        [OperationContract]
        [WebInvoke(UriTemplate = "/InsertInvoiceDetail",
        Method = "POST", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        InvoiceResponse InsertInvoiceDetail(InvoiceDelailInfo _data);

        [OperationContract]
        [WebInvoke(UriTemplate = "/SyncProductList?syncDateTime={syncDateTime}",
        Method = "GET", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        ModelDetailsList SyncProductList(string syncDateTime);

        [OperationContract]
        [WebInvoke(UriTemplate = "/SyncAMCPlanDetails?syncDateTime={syncDateTime}",
        Method = "GET", RequestFormat = WebMessageFormat.Json,
        ResponseFormat = WebMessageFormat.Json)]
        AMCPlanDetailsList SyncAMCPlanDetails(string syncDateTime);
    }
}