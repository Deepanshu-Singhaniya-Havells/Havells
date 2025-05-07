using ConsumerApp.BusinessLayer;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
#pragma warning disable CS0105 // The using directive for 'System' appeared previously in this namespace
#pragma warning restore CS0105 // The using directive for 'System' appeared previously in this namespace
#pragma warning disable CS0105 // The using directive for 'System.Linq' appeared previously in this namespace
#pragma warning restore CS0105 // The using directive for 'System.Linq' appeared previously in this namespace

namespace ConsumerAppServices
{
	// NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
	// NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
	public class ConsumerAppWS : IConsumerAppWS
	{
		public string GetData(int value)
		{
			return string.Format("You entered: {0}", value);
		}

		public List<JobOutput> GetJobs(Job InputJob)
		{
			return new Job().getJobs(InputJob);
		}
		public ForgotPassword ConsumerAppForgotPassword(ForgotPassword forgotPassword)
		{
			return new ForgotPassword().RunAction(forgotPassword);
		}
		public ForgotPasswordCheck1 ConsumerAppForgotPasswordCheck1(ForgotPasswordCheck1 forgotPassword)
		{
			return new ForgotPasswordCheck1().SendOTPFromService(forgotPassword);
		}
		public RegisterCustomer ConsumerAppRegisterCustomer(RegisterCustomer Register)
		{
			return new RegisterCustomer().InitiateCustomerCreation(Register);
		}
		public AuthenticateCustomer ConsumerAppAuthenticateCustomer(AuthenticateCustomer Validate)
		{
			return new AuthenticateCustomer().ValidateThisCustomer(Validate);
		}
		public List<CityMaster> ConsumerAppCityMaster(CityMaster City)
		{
			List<CityMaster> obj = new List<CityMaster>();
			obj = City.GetAllActiveCities(City);
			return obj;
		}
		public List<PinCodeMaster> ConsumerAppPincodeMaster(PinCodeMaster Pin)
		{
			List<PinCodeMaster> obj = new List<PinCodeMaster>();
			obj = Pin.GetAllActivePinCodes(Pin);
			return obj;
		}
		public List<StateMaster> ConsumerAppStateMaster()
		{
			List<StateMaster> obj = new List<StateMaster>();
			StateMaster St = new StateMaster();
			obj = St.GetAllActiveStateCodes();
			return obj;
		}
		public List<PinCodeMasterData> ConsumerAppPinCodeMaster()
		{
			List<PinCodeMasterData> obj = new List<PinCodeMasterData>();
			PinCodeMasterData St = new PinCodeMasterData();
			obj = St.GetAllActivePinCodesMaster();
			return obj;
		}
		public List<CallSubtypeMaster> ConsumerAppCallSubTypeMaster()
		{
			List<CallSubtypeMaster> obj = new List<CallSubtypeMaster>();
			CallSubtypeMaster CallSt = new CallSubtypeMaster();
			obj = CallSt.GetAllCallSubType();
			return obj;
		}
		public List<CallSubtypeMasterForService> CallSubtypeMasterService()
		{
			List<CallSubtypeMasterForService> obj = new List<CallSubtypeMasterForService>();
			CallSubtypeMasterForService Serv = new CallSubtypeMasterForService();
			obj = Serv.GetAllCallSubTypeService();
			return obj;
		}
		public List<CityUnFiltered> ConsumerAppCityUnfilter()
		{
			List<CityUnFiltered> obj1 = new List<CityUnFiltered>();
			CityUnFiltered Serv1 = new CityUnFiltered();
			obj1 = Serv1.GetAllActiveCitiesUnFiltered();
			return obj1;
		}
		public ChangePassword ConsumerAppChangePassword(ChangePassword details)
		{
			return new ChangePassword().SetNewPassword(details);
		}
		public ProfileUpdate ConsumerAppProfileUpdate(ProfileUpdate details)
		{
			return new ProfileUpdate().UpdateProfile(details);
		}
		public ProductRegistration ConsumerAppProductRegistration(ProductRegistration Register)
		{
			return new ProductRegistration().InitiateProductRegistration(Register);
		}
		public List<CustomerWishList> ConsumerAppCustomerWishList(CustomerWishList customerWishList)
		{
			List<CustomerWishList> obj = new List<CustomerWishList>();
			obj = customerWishList.GetAllCustomerWishList(customerWishList);
			return obj;
		}
		public CreateWorkOrder ConsumerAppCreateWorkOrder(CreateWorkOrder bridge)
		{
			return new CreateWorkOrder().SubmitServiceRequest(bridge);
		}
		public CreateWorkOrderQA CreateJobsinQA(CreateWorkOrderQA iMaster)
		{
			return new CreateWorkOrderQA().SubmitServiceRequestQA(iMaster);
		}
		public List<ProductCategorySubCategoryMaster> ConsumerAppProductCategorySubCategory()
		{
			List<ProductCategorySubCategoryMaster> obj = new List<ProductCategorySubCategoryMaster>();
			ProductCategorySubCategoryMaster Pdt = new ProductCategorySubCategoryMaster();
			obj = Pdt.GetProductCategory();
			return obj;
		}
		public List<ExistingServices> ConsumerAppExistingServices(ExistingServices Serv)
		{
			List<ExistingServices> obj = new List<ExistingServices>();
			obj = Serv.GetAllExistingServices(Serv);
			return obj;
		}
		public List<AlertsAndNotifications> ConsumerAppAlertsAndNotifications(AlertsAndNotifications bridge)
		{
			List<AlertsAndNotifications> obj = new List<AlertsAndNotifications>();
			obj = bridge.GetAlertsOnDate(bridge);
			return obj;
		}
		public List<AlertsAndNotificationsPerDay> ConsumerAppAlertsAndNotificationsPerDay(AlertsAndNotificationsPerDay bridge)
		{
			List<AlertsAndNotificationsPerDay> obj = new List<AlertsAndNotificationsPerDay>();
			obj = bridge.GetAlertsPerDay(bridge);
			return obj;
		}

		public List<ExisitingPurchases> ConsumerAppExistingPurchases(ExisitingPurchases bridge)
		{
			List<ExisitingPurchases> obj = new List<ExisitingPurchases>();
			obj = bridge.GetAllExistingPurchase(bridge);
			return obj;
		}

		public List<GeoLocator> ConsumerAppGeoLocator(GeoLocator bridge)
		{
			List<GeoLocator> obj = new List<GeoLocator>();
			obj = bridge.LocateAccounts(bridge);
			return obj;
		}
		public CreateWorkOrderForDemo ConsumerAppCreateWorkOrderForDemo(CreateWorkOrderForDemo bridge)
		{
			return new CreateWorkOrderForDemo().SubmitInstallationDemoRequest(bridge);
		}
		public ProfileView ConsumerAppCustomerProfileDetails(ProfileView bridge)
		{
			return new ProfileView().GetCustomerInformation(bridge);
		}
		public GetCustomerGuid ConsumerAppGetCustomerGuid(GetCustomerGuid bridge)
		{
			return new GetCustomerGuid().GetGuIdBasisMobNo(bridge);
		}
		public ConsumerAddress ConsumerAppAddressCreate(ConsumerAddress Addr)
		{
			return new ConsumerAddress().AddAddress(Addr);
		}
		public List<PinCodeFilter> ConsumerAppPincodeFilter(PinCodeFilter Pin)
		{
			List<PinCodeFilter> obj = new List<PinCodeFilter>();
			obj = Pin.GetAllActivePinCodes(Pin);
			return obj;
		}
		public IN_SMS InComingSMSCreate(string FROM, string TO, string MESSAGE, int JobType)
		{
			return new IN_SMS().CreateInComingSMS(FROM, TO, MESSAGE, JobType);
		}
		public PUSH_LEAD CreateLeadAvaya(PUSH_LEAD iLead)
		{
			return new PUSH_LEAD().CreateLeadForPreSalesEnquiry(iLead);
		}
		public GetCustomerName AVAYAGetCustomerFullName(GetCustomerName Cust)
		{
			return new GetCustomerName().GetFullNameBasisMobNo(Cust);
		}
		public GetLastOpenCase AVAYAGetLastOpenCase(GetLastOpenCase Cust)
		{
			return new GetLastOpenCase().GetLastOpenCaseBasisMobNo(Cust);
		}
		public EscalateWorkOrder AVAYAEscallateWrkOrder(EscalateWorkOrder Cust)
		{
			return new EscalateWorkOrder().EscallateWorkOd(Cust);
		}
		public AVAYACreateWo AVAYACreateWO(AVAYACreateWo Cust)
		{
			return new AVAYACreateWo().AvayaCreateWrkOrder(Cust);
		}
		public ReturnInfo ECommercePushCust(PUSH_Customer Cust)
		{
			ReturnInfo _retInfo = new ReturnInfo();
			_retInfo = Cust.InitiateOperation(Cust);
			return _retInfo;
		}
		public CreateContactFromBridge ConsumerAppValidatedOTP(CreateContactFromBridge crBridge)
		{
			return new CreateContactFromBridge().iCreateContactFromBridge(crBridge);
		}
		public JobDetails FindJobDetails(JobDetails iJob)
		{
			return new JobDetails().iGetJobDetails(iJob);
		}
		public List<ProductSubCategoryMaster> GetProductSubCategory(ProductSubCategoryMaster iCatMaster)
		{
			List<ProductSubCategoryMaster> obj = new List<ProductSubCategoryMaster>();
			obj = iCatMaster.GetProductSubCategory(iCatMaster);
			return obj;
		}
		public List<ProductCodeMaster> GetProductMaster(ProductCodeMaster iSCatMaster)
		{
			List<ProductCodeMaster> obj = new List<ProductCodeMaster>();
			obj = iSCatMaster.GetProductMaster(iSCatMaster);
			return obj;
		}
		public OUTGOING_SMS SendSMS(OUTGOING_SMS iDetails)
		{
			return new OUTGOING_SMS().OUTGOINGSMSMETHOD(iDetails);
		}

		public AuthenticationServiceExternal Authenticate(AuthenticationServiceExternal oSFA)
		{
			return new AuthenticationServiceExternal().Authenticate(oSFA);
		}

		public AuthenticateCustomerOTP iAuthenticate(AuthenticateCustomerOTP iDetails)
		{
			return new AuthenticateCustomerOTP().ValidateThisCustomerOTP(iDetails);
		}
		public Feedback iFeedback(Feedback iDetails)
		{
			//return new Feedback().PostFeedback(iDetails);
			return null;
		}
		public YoutubeApi iYoutube()
		{
			return new YoutubeApi().ReturnYoutubeConfigs();
		}
		public FindMaterialByName GetMaterialDetailsbyName(FindMaterialByName iDetails)
		{
			return new FindMaterialByName().GetMaterialId(iDetails);
		}
		public ValidateSerialNumber iValidateSerialNumFromSAP(ValidateSerialNumber iValidate)
		{
			return new ValidateSerialNumber().GetResponseFromSAP(iValidate);
		}


		public IsExistingUser UserValidation(IsExistingUser iCheck)
		{
			return new IsExistingUser().CheckIfDuplicate(iCheck);
		}
		public bool FeatureInMaking()
		{
			return IsFeatureAvailable.ReturnIfFeatureNotAvailable();
		}
		public List<OutputAddressClass> GetAllUserAddress(GetUserAddresses iConsumer)
		{
			GetUserAddresses iUser = new GetUserAddresses();
			List<OutputAddressClass> obj = new List<OutputAddressClass>();
			obj = iUser.GetCustomerInformation(iConsumer);
			return obj;
		}
		public AddressFields ViewThisAddress(ViewAddress iViewAddress)
		{
			ViewAddress iView = new ViewAddress();
			AddressFields iAddress = new AddressFields();
			iAddress = iView.GetThisAddress(iViewAddress);
			return iAddress;
		}
		public OutputAddAddress AddConsumerAddress(AddAddress iAdd)
		{
			OutputAddAddress iOutput = new OutputAddAddress();
			AddAddress iAddAddress = new AddAddress();
			iOutput = iAddAddress.AddCustomerAddress(iAdd);
			return iOutput;
		}
		public OutputUpdateAddress UpdateConsumerAddress(UpdateAddress iUpdate)
		{
			OutputUpdateAddress iOutput = new OutputUpdateAddress();
			UpdateAddress iAddAddress = new UpdateAddress();
			iOutput = iAddAddress.UpdateThisAddress(iUpdate);
			return iOutput;
		}
		public OutputCreateServiceTicket_Breakdown CreateServiceTicketNew(CreateServiceTicket_Breakdown iCreateTicket)
		{
			OutputCreateServiceTicket_Breakdown iOutput = new OutputCreateServiceTicket_Breakdown();
			CreateServiceTicket_Breakdown iInput = new CreateServiceTicket_Breakdown();
			iOutput = iInput.SubmitJob_New(iCreateTicket);
			return iOutput;
		}
		public OutputAddressDelete DeleteCustomerAddress(AddressDelete iDeleteAdd)
		{
			OutputAddressDelete iOut = new OutputAddressDelete();
			AddressDelete iInput = new AddressDelete();
			iOut = iInput.DeleteThisAddress(iDeleteAdd);
			return iOut;
		}
		public ValidateSerialNumberQAcs iValidateSerialNumFromSAPQA(ValidateSerialNumberQAcs iValidate)
		{
			return new ValidateSerialNumberQAcs().GetResponseFromSAPQA(iValidate);
		}
		public List<OutputDistrictMaster> GetDistrictMaster(DistrictMaster iMaster)
		{
			List<OutputDistrictMaster> iOut = new List<OutputDistrictMaster>();
			DistrictMaster iInput = new DistrictMaster();
			iOut = iInput.GetListofDistricts(iMaster);
			return iOut;
		}
		public ReturnInfoVar CreateEnquiryEntry(CreateEnquiry enquiry)
		{
			return enquiry.CreateEnquiryEntry(enquiry);
		}

		public TechnicianInfo GetTechnicianInfo(TechnicianInfo care360ID)
		{
			TechnicianInfo _retInfo = new TechnicianInfo();
			_retInfo = _retInfo.GetTechnicianInfo(care360ID);
			return _retInfo;
		}

		public List<RegisteredProducts> GetRegisteredProducts(RegisteredProducts registeredProducts)
		{
			List<RegisteredProducts> _retInfo = new List<RegisteredProducts>();
			RegisteredProducts _registeredProductInfo = new RegisteredProducts();
			_retInfo = _registeredProductInfo.GetRegisteredProducts(registeredProducts);
			return _retInfo;
		}


		public List<NatureOfComplaint> GetNatureOfComplaints(NatureOfComplaint natureOfComplaint)
		{
			List<NatureOfComplaint> _retInfo = new List<NatureOfComplaint>();
			NatureOfComplaint _registeredProductInfo = new NatureOfComplaint();
			_retInfo = _registeredProductInfo.GetNatureOfComplaints(natureOfComplaint);
			return _retInfo;
		}

		public List<NatureOfComplaint> GetAllNOCs()
		{
			List<NatureOfComplaint> _retInfo = new List<NatureOfComplaint>();
			NatureOfComplaint _registeredProductInfo = new NatureOfComplaint();
			_retInfo = _registeredProductInfo.GetAllNOCs();
			return _retInfo;
		}

		public List<AddressBook> GetAddress(AddressBook natureOfComplaint)
		{
			List<AddressBook> _retInfo = new List<AddressBook>();
			AddressBook _addressBookInfo = new AddressBook();
			_retInfo = _addressBookInfo.GetAddresses(natureOfComplaint);
			return _retInfo;
		}

		public ServiceCall CreateServiceCall(ServiceCall ServiceCallData)
		{
			ServiceCall _retInfo = new ServiceCall();
			return _retInfo.CreateServiceCall(ServiceCallData);
		}

		#region Havells IoT Platform

		#region Consumer APIs
		public ReturnResult RegisterConsumer(IoT_RegisterConsumer consumer)
		{
			ReturnResult _retInfo = new ReturnResult();
			_retInfo = consumer.RegisterConsumer(consumer);
			return _retInfo;
		}
		public List<HashTableDTO> GetSalutationEnum()
		{
			List<HashTableDTO> _retInfo = new List<HashTableDTO>();
			IoTCommonLib _commonLib = new IoTCommonLib();
			_retInfo = _commonLib.GetSalutationEnum();
			return _retInfo;
		}
		public IoTCustomerProfile GetIoTCustomerProfile(IoTCustomerProfile customerProfileData)
		{
			IoTCustomerProfile _customerProfileInfo = new IoTCustomerProfile();
			_customerProfileInfo = _customerProfileInfo.GetIoTCustomerProfile(customerProfileData);
			return _customerProfileInfo;
		}
		public IoTCustomerProfile UpdateIoTCustomerProfile(IoTCustomerProfile customerProfileData)
		{
			IoTCustomerProfile _customerProfileInfo = new IoTCustomerProfile();
			_customerProfileInfo = _customerProfileInfo.UpdateIoTCustomerProfile(customerProfileData);
			return _customerProfileInfo;
		}
		#endregion

		#region Address Book APIs
		public List<IoTAddressBookResult> GetIoTAddressBook(IoTAddressBook address)
		{
			List<IoTAddressBookResult> _retInfo = new List<IoTAddressBookResult>();
			IoTAddressBook _addressBookInfo = new IoTAddressBook();
			_retInfo = _addressBookInfo.GetIoTAddressBook(address);
			return _retInfo;
		}
		public IoTPINCodes IoTValidatePINCode(IoTPINCodes address)
		{
			IoTPINCodes _pinCodeInfo = new IoTPINCodes();
			IoTAddressBook _addressInfo = new IoTAddressBook();
			_pinCodeInfo = _addressInfo.IoTValidatePINCode(address);
			return _pinCodeInfo;
		}
		public List<IoTAreas> IoTGetAreasByPinCode(IoTPINCodes address)
		{
			List<IoTAreas> _areaInfo = new List<IoTAreas>();
			IoTAddressBook _addressInfo = new IoTAddressBook();
			_areaInfo = _addressInfo.IoTGetAreasByPinCode(address);
			return _areaInfo;
		}
		public IoTAddressBookResult IoTCreateAddress(IoTAddressBookResult address)
		{
			IoTAddressBook _addressInfo = new IoTAddressBook();
			IoTAddressBookResult _addressResultInfo = new IoTAddressBookResult();
			_addressResultInfo = _addressInfo.IoTCreateAddress(address);
			return _addressResultInfo;
		}
		public IoTAddressBookResult IoTUpdateAddress(IoTAddressBookResult address)
		{
			IoTAddressBook _addressInfo = new IoTAddressBook();
			IoTAddressBookResult _addressResultInfo = new IoTAddressBookResult();
			_addressResultInfo = _addressInfo.IoTUpdateAddress(address);
			return _addressResultInfo;
		}
		public IoTAddressBookResult ECommerceBulkCustomerData(IoTAddressBookResult addressData)
		{
			IoTAddressBook _addressInfo = new IoTAddressBook();
			IoTAddressBookResult _addressResultInfo = new IoTAddressBookResult();
			_addressResultInfo = _addressInfo.ECommerceBulkCustomerData(addressData);
			return _addressResultInfo;
		}

		public List<IoTAddressBookResultV1> GetIoTAddressBookV1(IoTAddressBookV1 address)
		{
			List<IoTAddressBookResultV1> _retInfo = new List<IoTAddressBookResultV1>();
			IoTAddressBookV1 _addressBookInfo = new IoTAddressBookV1();
			_retInfo = _addressBookInfo.GetIoTAddressBook(address);
			return _retInfo;
		}

		public IoTAddressBookResultV1 IoTCreateAddressV1(IoTAddressBookResultV1 address)
		{
			IoTAddressBookV1 _addressInfo = new IoTAddressBookV1();
			IoTAddressBookResultV1 _addressResultInfo = new IoTAddressBookResultV1();
			_addressResultInfo = _addressInfo.IoTCreateAddress(address);
			return _addressResultInfo;
		}

		public IoTAddressBookResultV1 IoTUpdateAddressV1(IoTAddressBookResultV1 address)
		{
			IoTAddressBookV1 _addressInfo = new IoTAddressBookV1();
			IoTAddressBookResultV1 _addressResultInfo = new IoTAddressBookResultV1();
			_addressResultInfo = _addressInfo.IoTUpdateAddress(address);
			return _addressResultInfo;
		}
		public IoTAddressBookResultV1 IoTDeleteAddressV1(IoTAddressBookResultV1 address)
		{
			IoTAddressBookV1 _addressInfo = new IoTAddressBookV1();
			IoTAddressBookResultV1 _addressResultInfo = new IoTAddressBookResultV1();
			_addressResultInfo = _addressInfo.IoTDeleteAddress(address);
			return _addressResultInfo;
		}

		#endregion

		#region Consumer Products APIs
		public List<IoTRegisteredProducts> GetIoTRegisteredProducts(IoTRegisteredProducts registeredProducts)
		{
			List<IoTRegisteredProducts> _retInfo = new List<IoTRegisteredProducts>();
			IoTRegisteredProducts _registeredProductInfo = new IoTRegisteredProducts();
			_retInfo = _registeredProductInfo.GetIoTRegisteredProducts(registeredProducts);
			return _retInfo;
		}

		public List<IoTRegisteredProducts> GetIoTRegisteredProductsProd(IoTRegisteredProducts registeredProducts)
		{
			List<IoTRegisteredProducts> _retInfo = new List<IoTRegisteredProducts>();
			IoTRegisteredProducts _registeredProductInfo = new IoTRegisteredProducts();
			_retInfo = _registeredProductInfo.GetIoTRegisteredProductsProd(registeredProducts);
			return _retInfo;
		}

		public List<HashTableDTO> GetInstallationLocationEnum()
		{
			List<HashTableDTO> _retInfo = new List<HashTableDTO>();
			IoTCommonLib _commonLib = new IoTCommonLib();
			_retInfo = _commonLib.GetInstallationLocationEnum();
			return _retInfo;
		}
		public List<ProductHierarchyDTO> GetProductHierarchy()
		{
			List<ProductHierarchyDTO> _retInfo = new List<ProductHierarchyDTO>();
			IoTCommonLib _commonLib = new IoTCommonLib();
			_retInfo = _commonLib.GetProductHierarchy();
			return _retInfo;
		}
		public List<ProductDTO> GetProducts(IoTCommonLib ProductSubCatgInfo)
		{
			List<ProductDTO> _retInfo = new List<ProductDTO>();
			IoTCommonLib _commonLib = new IoTCommonLib();
			_retInfo = _commonLib.GetProducts(ProductSubCatgInfo);
			return _retInfo;
		}
		public Attachment AttachNotes(Attachment attachmentData)
		{
			Attachment _retObj = new Attachment();
			IoTCommonLib commonLib = new IoTCommonLib();
			_retObj = commonLib.AttachNotes(attachmentData);
			return _retObj;
		}
		public IoTValidateSerialNumber ValidateAssetSerialNumber(IoTValidateSerialNumber _reqParam)
		{
			IoTValidateSerialNumber _retInfo = new IoTValidateSerialNumber();
			return _retInfo.ValidateAssetSerialNumber(_reqParam);
		}
		public IoTRegisteredProductsResult RegisterProduct(IoTRegisteredProducts productData)
		{
			return new IoTRegisteredProducts().RegisterProduct(productData);
		}
		public IoTValidateSerialNumber GetProductDetails(IoTValidateSerialNumber _reqParam)
		{
			return new IoTValidateSerialNumber().GetProductDetails(_reqParam);
		}
		#endregion

		#region Service Call APIs
		public List<IoTServiceCallResult> GetIoTServiceCalls(IotServiceCall job)
		{
			List<IoTServiceCallResult> _retInfo = new List<IoTServiceCallResult>();
			IotServiceCall _serviceCallInfo = new IotServiceCall();
			_retInfo = _serviceCallInfo.GetIoTServiceCalls(job);
			return _retInfo;
		}

		public IoTServiceCallRegistration IoTCreateServiceCall(IoTServiceCallRegistration serviceCalldata)
		{
			IotServiceCall serviceCall = new IotServiceCall();
			IoTServiceCallRegistration serviceCallRegistrationResult = new IoTServiceCallRegistration();
			serviceCallRegistrationResult = serviceCall.IoTCreateServiceCall(serviceCalldata);
			return serviceCallRegistrationResult;
		}

		public List<IoTNatureofComplaint> GetIoTNatureOfComplaints(IoTNatureofComplaint natureOfComplaint)
		{
			IotServiceCall serviceCall = new IotServiceCall();
			List<IoTNatureofComplaint> _retInfo = new List<IoTNatureofComplaint>();
			_retInfo = serviceCall.GetIoTNatureOfComplaints(natureOfComplaint);
			return _retInfo;
		}

		public List<IoTNatureofComplaint> GetIoTNatureOfComplaintsByProdSubCatg(IoTNatureofComplaint natureOfComplaint)
		{
			IotServiceCall serviceCall = new IotServiceCall();
			List<IoTNatureofComplaint> _retInfo = new List<IoTNatureofComplaint>();
			_retInfo = serviceCall.GetIoTNatureOfComplaintsByProdSubCatg(natureOfComplaint);
			return _retInfo;
		}

		public IoTServiceCallRegistration IoTCreateServiceCallWhatsapp(IoTServiceCallRegistration serviceCalldata)
		{
			IotServiceCall serviceCall = new IotServiceCall();
			IoTServiceCallRegistration serviceCallRegistrationResult = new IoTServiceCallRegistration();
			serviceCallRegistrationResult = serviceCall.IoTCreateServiceCallWhatsapp(serviceCalldata);
			return serviceCallRegistrationResult;
		}
		#endregion

		#endregion

		public ValidateKKGCodeResult ValidateKKGCode(KKGCodeValidator jobData)
		{
			ValidateKKGCodeResult retObj = new ValidateKKGCodeResult();
			KKGCodeValidator validatorObj = new KKGCodeValidator();
			retObj = validatorObj.ValidateKKGCode(jobData);
			return retObj;
		}
		public IoTRegisteredProductsResult RegisterProductFromWhatsapp(IoTRegisteredProducts productData)
		{
			IoTRegisteredProductsResult retObj = new IoTRegisteredProductsResult();
			IoTRegisteredProducts CustAssetObj = new IoTRegisteredProducts();
			retObj = CustAssetObj.RegisterProductFromWhatsapp(productData);
			return retObj;
		}

		public FlushUATDataResult FlushData(FlushUATData flushUATData)
		{
			FlushUATDataResult retObj = new FlushUATDataResult();
			FlushUATData flushDataObj = new FlushUATData();
			retObj = flushDataObj.FlushData(flushUATData);
			return retObj;
		}

		public UserManagement UpdateBusinessUnit(UserManagement userData)
		{
			UserManagement retObj = new UserManagement();
			UserManagement buDataObj = new UserManagement();
			retObj = buDataObj.UpdateBusinessUnit(userData);
			return retObj;
		}

		public ValidateSerialNumber ValidateSerialNumberThirdParty(ValidateSerialNumber iValidate)
		{
			return new ValidateSerialNumber().ValidateSerialNumberWithSAP(iValidate);
		}

		#region SFA Module APIs
		public SFA_ValidateCustomer ValidateCustomer(SFA_ValidateCustomer customerData)
		{
			SFA_ValidateCustomer retObj = new SFA_ValidateCustomer();
			return retObj.ValidateCustomer(customerData);
		}
		public SFA_ServiceCallResult SFA_CreateServiceCall(SFA_ServiceCall serviceCallData)
		{
			SFA_ServiceCall retObj = new SFA_ServiceCall();
			return retObj.SFA_CreateServiceCall(serviceCallData);
		}
		public List<SFA_DivisionCallType> GetDivisionCallTypeSetup()
		{
			SFA_DivisionCallType retObj = new SFA_DivisionCallType();
			return retObj.GetDivisionCallTypeSetup();
		}
		#endregion

		#region D365 Audit Log
		public List<AttributeMetadataInfo> AttributeMetadata(EntityMetadataInfo entity)
		{
			EntityMetadataInfo retObj = new EntityMetadataInfo();
			return retObj.AttributeMetadata(entity);
		}

		public List<D365AuditLogResult> GetD365AuditLogData(D365AuditLog _requestData)
		{
			D365AuditLog retObj = new D365AuditLog();
			return retObj.GetD365AuditLogData(_requestData);
		}
		#endregion

		#region MRN Module APIs
		public List<CustomerInfo> GetCustomerDetails(string CustomerSearchString)
		{
			MRNEntry retObj = new MRNEntry();
			return retObj.GetCustomerDetails(CustomerSearchString);
		}
		public ValidateSerialNumberDEV GetResponseFromSAPDEV(ValidateSerialNumberDEV iValidate)
		{
			ValidateSerialNumberDEV _retObj = new ValidateSerialNumberDEV();
			return _retObj.GetResponseFromSAPDEV(iValidate);
		}

		public MRNHeader AddtoViewList(MRNEntry _productInfo)
		{
			return new MRNEntry().AddtoViewList(_productInfo);
		}
		public MRNHeader SubmitViewList(MRNHeader _viewList)
		{
			return new MRNEntry().SubmitViewList(_viewList);
		}

		public MRNSummary GetViewList(ViewListSearch _searchCondition)
		{
			return new MRNEntry().GetViewList(_searchCondition);
		}
		public List<PDOICalls> GetPDICalls(MRNEntry job)
		{
			return new MRNEntry().GetPDICalls(job);
		}

		public MRNSummary GetDefectiveStockNoteForSAP(ViewListSearch _searchCondition)
		{
			return new MRNEntry().GetDefectiveStockNoteForSAP(_searchCondition);
		}

		#endregion

		#region Tender Module
		public List<DocumentTypes> GetDocumentTypes()
		{
			return new Tender().GetDocumentTypes();
		}

		#endregion

		#region Attachment Manager
		public List<DocumemtType> AMGetDocumentTypes(DocumemtType docType)
		{
			return new AttachmentManager().GetDocumentTypes(docType);
		}

		#endregion

		#region Claim Extension on Tech Mobile
		public List<WOSchemesResult> GetSchemeCodes(ClaimExtToMobile _reqParam)
		{
			return new ClaimExtToMobile().GetSchemeCodes(_reqParam);
		}
		#endregion

		#region AMC Billing
		public AMCBilling ValidateAMCReceiptAmount(AMCBilling _reqData)
		{
			return new AMCBilling().ValidateAMCReceiptAmount(_reqData);
		}
		#endregion

		#region Remote Pay
		public SendPaymentUrlResponse SendRemotePaySMS(SendURLD365Request reqParm)
		{
			return new RemotePay().SendSMS(reqParm);
		}
		public PaymentStatusD365Response getRemotePayStatus(String jobId)
		{
			return new RemotePay().getPaymentStatus(jobId);
		}
		#endregion

		#region Home Advisory
		public Response CreateAppointment(CRMRequest req)
		{
			return new HomeAdvisory().CreateAppointmentD365(req);
		}
		public HopmeAdvisoryResult CreateAdvisory(HomeAdvisoryRequest reqParm)
		{
			return new HomeAdvisory().CreateAdvisory(reqParm);
		}

		public GetEnquiry GetEnquery(Enquiry req)
		{
			return new HomeAdvisory().GetEnquery(req);
		}

		public GetEnquiry GetEnqueryStatus(EnquiryStatus req)
		{
			return new HomeAdvisory().GetEnqueryStatus(req);
		}

		public List<EnquiryDetails> GetSalesEnquiry(EnquiryStatus _enquiryStatus)
		{
			return new CreateEnquiry().GetSalesEnquiry(_enquiryStatus);
		}
		public Response RescheduleAppointment(ReschduleAppointment req)
		{
			return new HomeAdvisory().RescheduleAppointment(req);
		}
		public Response CancleAppointment(CancelAppointmentRequest req)
		{
			//CancelAppointmentRequest req = JsonConvert.DeserializeObject<CancelAppointmentRequest>(SecurityUtility.Decrypt(reqJSON));
			//return new HomeAdvisory().CancleAppointment(req);
			return new HomeAdvisory().CancleAppointment(req);
		}
		public ResponseUpload UploadAttachment(UploadAttachment reqParm)
		{
			return new HomeAdvisory().UploadAttachment(reqParm);
		}

		public GetUserTimeSlotsRoot GetUserTimeSlots(GetUserTimeSlotsRequest requestParm)
		{
			return new HomeAdvisory().GetUserTimeSlots(requestParm);
		}

		public List<EnquiryType> GetEnquiryType(EnquiryType _enquiryCategory)
		{
			return new CreateEnquiry().GetEnquiryTypes(_enquiryCategory);
		}

		public List<EnquiryProductType> GetEnquiryProductType(EnquiryProductType _enquiryCategory)
		{
			return new CreateEnquiry().GetEnquiryProductTypes(_enquiryCategory);
		}

		#endregion

		#region Experience Store
		public ExperinceZonePayLoad ESRegisterConsumer(ExperinceZonePayLoad consumer)
		{
			return new ExperinceZoneAPI().RegisterConsumer(consumer);
		}
		#endregion

		#region Consumer Survey NPS
		public ConsumerSurveyDTO GetConsumerSurvey(ConsumerSurveyDTO _consumerSurveyData)
		{
			return new ConsumerNPS().GetConsumerSurvey(_consumerSurveyData);
		}
		public ConsumerSurveyDTO CaptureConsumerSurvey(ConsumerSurveyDTO _consumerSurveyData)
		{
			return new ConsumerNPS().CaptureConsumerSurvey(_consumerSurveyData);
		}
		#endregion

		#region APIs to retreive D365 Archived Data
		public ArchivedJobsResponse GetArchivedJobData(ArchivedJobRequest inpReq)
		{
			return new D365ArchivedData().GetJobData(inpReq);
		}
		public ResponseData GetArchivedData(RequestData inpReq)
		{
			return new D365ArchivedData().GetArchiveData(inpReq);
		}

		#endregion

		#region Dealer Portal
		public IoTServiceCallRegistration IoTCreateServiceCallDealerPortal(IoTServiceCallRegistration serviceCalldata)
		{
			return new IotServiceCall().IoTCreateServiceCallDealerPortal(serviceCalldata);
		}
		#endregion

		#region D365 odata Access Token
		public string getAccessToken()
		{
			return new AccessToken().getAccessToken();
		}
		#endregion

		public List<Makes> GetMakes()
		{
			return new DemoAPIs().GetMakes();
		}

		#region Consumer NPS
		public ResponseDTO SendCommunication(RequestDTO _data)
		{
			return new CommunicationNPS().SendCommunication(_data);
		}
		#endregion
		public CancelJobResponse CancelServiceJob(CancelJobRequest reqParam)
		{
			return new IotServiceCall().CancelServiceJob(reqParam);
		}

		#region Voice Bot APIs
		public BotSpeechTransScript InsertBotSpeechTransScript(BotSpeechTransScript reqParams)
		{
			return new BotSpeechTransScript().InsertBotSpeechTransScript(reqParams);
		}
		public CallbackRequest InsertCallbackRequest(CallbackRequest reqParams)
		{
			return new CallbackRequest().InsertCallbackRequest(reqParams);
		}
		public Escalations GetEscalations(Escalations JobData)
		{
			return new Escalations().GetEscalations(JobData);
		}
		public Escalations InsertEscalations(EscalationsReqRes reqParams)
		{
			return new Escalations().InsertEscalations(reqParams);
		}
		public ProductDivision GetProductDivisions()
		{
			return new ProductDivision().GetProductDivisions();
		}
		public IoTServiceCallRegistration IoTCreateServiceVoiceBot(IoTServiceCallRegistration serviceCalldata)
		{
			return new IotServiceCall().IoTCreateServiceVoiceBot(serviceCalldata);
		}
		public IoTServiceCallRegistration IoTCreateServiceVoiceBotUAT(IoTServiceCallRegistration serviceCalldata)
		{
			return new IotServiceCall().IoTCreateServiceVoiceBotUAT(serviceCalldata);
		}
		#endregion
		#region Claim Performa
		public AnnextureResponse GenerateAnnexureUrl(string _PerformaInvoice)
		{
			return new ClaimPerforma().GenerateAnnexureUrl(_PerformaInvoice);
		}

		#endregion
		public ValidateKKGCodeResult KKGCodeVerification(ValidateKKGCodeInput validateKKGCodeInput)
		{
			return new KKGCodeValidator().KKGCodeVerification(validateKKGCodeInput);
		}

		public ServiceResponseData GetOutstandingAMCs(ReqestData reqestData)
		{
			return new AMCBilling().GetOutstandingAMCs(reqestData);
		}

		public AuthenticateConsumer AuthenticateConsumerAMC(AuthenticateConsumer requestParam)
		{
			return new IotServiceCall().AuthenticateConsumerAMC(requestParam);
		}
		public ValidateSessionResponse ValidateSessionDetails(ValidateSessionRequest requestParam)
		{
			return new IotServiceCall().ValidateSessionDetails(requestParam);
		}
		public ValidateSessionResponse CreateSession(AuthenticateConsumer requestParam)
		{
			return new IotServiceCall().CreateSession(requestParam);
		}
		public AuthModel AuthenticateSamparkAppLogin(AuthModel request_param)
		{
			return new SecurityUtility().AuthenticateSamparkAppLogin(request_param);
		}

        #region MFR Service Job
        public JobRequestDTO CreateServiceCallRequest(JobRequestDTO _jobRequest)
		{
			return new MFRServiceJobs().CreateServiceCallRequest(_jobRequest);
		}
		public JobStatusDTO GetJobstatus(JobStatusDTO _jobRequest)
		{
			return new MFRServiceJobs().GetJobstatus(_jobRequest);
		}

        public WorkOrderResponse GetWorkOrdersStatus(WorkOrderRequest objreq)
        {
            return new MFRServiceJobs().GetWorkOrdersStatus(objreq);
        }

        #endregion
        public GlobalSearchResponse GetGlobalSearch(GlobalSearchRequest req)
		{
			return new D365Metadata().GetGlobalSearch(req);
		}

		#region Consumable's History
		public ConsumablesData RequestData(Request req)
		{
			return new TechnicianMobileExt().RequestData(req);
		}
		#endregion

		public ValidateProductInstallation ValidateProductInstallation(ValidateProductInstallation _data)
		{
			return new IotServiceCall().ValidateProductInstallation(_data);
		}

		#region Airtel
		public ResposeDataCallMasking GetCustomerOpenJobs(RequestDataCallMasking _requestData)
		{
			return new AirtelIQ().GetOpenJobs(_requestData);
		}
		public CDR_Response PushCDR(CDR_Request request)
		{
			return new AirtelIQ().PushCDRToD365(request);
		}
		#endregion

		#region IDP Process
		public IDPProcessResponse UpdateIDPProcess(IDPProcessDetails request)
		{
			return new IDPProcess().UpdateIDPProcess(request);
		}

		public InvoiceResponse InsertInvoiceDetail(InvoiceDelailInfo request)
		{
			return new PushInvoiceDetails().InsertInvoiceDetail(request);
		}
		public ModelDetailsList SyncProductList(string syncDateTime)
		{
			return new PushInvoiceDetails().SyncProductList(syncDateTime);
		}
		public AMCPlanDetailsList SyncAMCPlanDetails(string syncDateTime)
		{
			return new PushInvoiceDetails().SyncAMCPlanDetails(syncDateTime);
		}
		#endregion

		#region Ecom Priceinfo
		public Ecom_PriceDetailsResponse GetEcom_Priceinfo(Ecom_PriceDetailsRequest request)
		{
			return new Ecom_Priceinfo().GetEcom_Priceinfo(request);
		}

		public BulkJobsModel CreateBulkJobs(BulkJobsModel BulkJobsData)
		{
			return new UploadBulkJob_RPA().CreateBulkJobs(BulkJobsData);
		}
		#endregion

		#region LitmusWorld
		public LtimusFeedBackResponse UpdateLitmusCustomerFeedBack(LtimusCustomerFeedBack ParamCustFeedBack)
		{
			LtimusFeedBackResponse objresult = new LtimusFeedBackResponse();
			objresult.Status = false;

			if (string.IsNullOrEmpty(ParamCustFeedBack.JobId))
			{
				objresult.Message = "Please enter jobid.";
				return objresult;
			}
			if (string.IsNullOrEmpty(ParamCustFeedBack.Category))
			{
				objresult.Message = "Please enter category.";
				return objresult;
			}
			if (ParamCustFeedBack.Score < 0 || ParamCustFeedBack.Score > 10)
			{
				objresult.Message = "Please enter valid score number between 1 to 10.";
				return objresult;
			}
			objresult = new LitmusFeedback().UpdateCustFeedBack(ParamCustFeedBack);
			return objresult;
		}


		#endregion
		#region Whatsapp Campain
		public KKGResponse UpdateKKGConsentOnJob(KKGRequest kKGRequest)
		{
			return new WhatsappCampaign().UpdateKKGConsentOnJob(kKGRequest);
		}
		#endregion
		#region Loyalty Program
		public sendEnrollmentRes SendEnrollmentLink(sendEnrollmentReq sendEnrollmentReq)
		{
			return new LoyaltyProgram().SendEnrollmentLink(sendEnrollmentReq);
		}

		#endregion
		#region Ala carte		
		public OrderCreateresponse CreateSalesOrder(CreateSalesOrderModel createSalesOrderModel)
		{
			return new ALaCarte().CreateSalesOrder(createSalesOrderModel);
		}

        #endregion
    }
}