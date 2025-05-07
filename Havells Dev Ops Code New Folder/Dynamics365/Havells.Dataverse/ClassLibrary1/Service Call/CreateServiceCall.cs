using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Havells.Dataverse.CustomConnector.Service_Call
{
    public class CreateServiceCall : IPlugin
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
            Guid CustomerGuid = Guid.Empty;
            Guid NOCGuid = Guid.Empty;
            Guid AddressGuid = Guid.Empty;
            Guid ProductCategoryGuid = Guid.Empty;
            Guid ProductSubCategoryGuid = Guid.Empty;
            StringBuilder errorMessage = new StringBuilder();
            bool IsValidRequest = true;
            bool isValid = true;            

            Regex Regex_MobileNo = new Regex("^[6-9]\\d{9}$");
            Regex Regex_PreferredDate = new Regex("(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)[0-9]{2}");
            if (context.InputParameters.Contains("CustomerMobleNo"))
            {
                string CustomerMobleNo = Convert.ToString(context.InputParameters["CustomerMobleNo"]);
                if (string.IsNullOrWhiteSpace(CustomerMobleNo))
                {
                    errorMessage.AppendLine("Customer mobile number is required.");
                    IsValidRequest = false;
                }
                else if (!Regex_MobileNo.IsMatch(CustomerMobleNo))
                {
                    errorMessage.AppendLine("Invalid customer mobile number.");
                    IsValidRequest = false;
                }
                string ChiefComplaint = Convert.ToString(context.InputParameters["ChiefComplaint"]);
                if (!string.IsNullOrWhiteSpace(ChiefComplaint))
                {
                    if (ChiefComplaint.Length > 2000)
                    {
                        errorMessage.AppendLine("Maximum length for Chief Complaint is 2000");
                        IsValidRequest = false;
                    }
                }
                int SourceOfJob = 0;
                int.TryParse(Convert.ToString(context.InputParameters["SourceOfJob"]), out SourceOfJob);

                if (SourceOfJob == 19 || SourceOfJob == 20) //19(MFR) and 20(Flipkart)
                {
                    string CustomerFirstName = Convert.ToString(context.InputParameters["CustomerFirstName"]);
                    string CustomerLastName = Convert.ToString(context.InputParameters["CustomerLastName"]);

                    if (string.IsNullOrWhiteSpace(CustomerFirstName))
                    {
                        errorMessage.AppendLine("Customer firstname is required.");
                        IsValidRequest = false;
                    }
                    else if (!APValidate.IsValidString(CustomerFirstName))
                    {
                        errorMessage.AppendLine("Invalid Customer First Name");
                        IsValidRequest = false;
                    }
                    else if (CustomerFirstName.Length > 100)
                    {
                        errorMessage.AppendLine("Maximum length for customer first name is 100");
                        IsValidRequest = false;
                    }
                    if (CustomerLastName.Length > 100)
                    {
                        errorMessage.AppendLine("Maximum length for customer last name is 100");
                        IsValidRequest = false;
                    }
                    string Addressline1 = Convert.ToString(context.InputParameters["AddressLine1"]);
                    string Addressline2 = Convert.ToString(context.InputParameters["AddressLine2"]);
                    string Landmark = Convert.ToString(context.InputParameters["Landmark"]);
                    string AlternateNumber = Convert.ToString(context.InputParameters["AlternateNumber"]);

                    if (string.IsNullOrWhiteSpace(Addressline1))
                    {
                        errorMessage.AppendLine("Customer Addressline1 is required.");
                        IsValidRequest = false;
                    }
                    else if (Addressline1.Length > 100)
                    {
                        errorMessage.AppendLine("Maximum length for Addressline1 is 100");
                        IsValidRequest = false;
                    }
                    if (!string.IsNullOrWhiteSpace(Addressline2))
                    {
                        if (Addressline2.Length > 100)
                        {
                            errorMessage.AppendLine("Maximum length for Addressline2 is 100");
                            IsValidRequest = false;
                        }
                    }
                    if (!string.IsNullOrWhiteSpace(AlternateNumber.Trim()))
                    {
                        if (AlternateNumber.Trim().Length > 10)
                        {
                            errorMessage.AppendLine("Maximum length for Alternate Number is 10");
                            IsValidRequest = false;
                        }
                        else if (!Regex_MobileNo.IsMatch(AlternateNumber.Trim()))
                        {
                            errorMessage.AppendLine("Invalid Alternate Number.");
                            IsValidRequest = false;
                        }
                    }
                    if (Landmark.Length > 100)
                    {
                        errorMessage.AppendLine("Maximum length for landmark is 100");
                        IsValidRequest = false;
                    }
                    string Pincode = Convert.ToString(context.InputParameters["Pincode"]);
                    if (string.IsNullOrWhiteSpace(Pincode))
                    {
                        errorMessage.AppendLine("Pincode is required.");
                        IsValidRequest = false;
                    }
                    else if (!Regex.IsMatch(Pincode, @"^\d{6}$"))
                    {
                        errorMessage.AppendLine("Invalid Pincode");
                        IsValidRequest = false;
                    }
                    string CallType = Convert.ToString(context.InputParameters["CallType"]);
                    if (string.IsNullOrWhiteSpace(CallType))
                    {
                        errorMessage.AppendLine("Call type is required.");
                        IsValidRequest = false;
                    }
                    else
                    {
                        string[] CallTypeArray = new string[] { "INSTALLATION", "BREAKDOWN" };
                        if (!CallTypeArray.Contains(CallType.ToUpper()))
                        {
                            errorMessage.AppendLine("Please Enter Correct Call Type.");
                            IsValidRequest = false;
                        }
                    }
                    string ProductSubCategoryName = Convert.ToString(context.InputParameters["ProductSubCategoryName"]);
                    if (string.IsNullOrWhiteSpace(ProductSubCategoryName))
                    {
                        errorMessage.AppendLine("Product Subcategory Name is required.");
                        IsValidRequest = false;
                    }
                    else if (ProductSubCategoryName.Length > 100)
                    {
                        errorMessage.AppendLine("Maximum length for Product Subcategory Name is 100");
                        IsValidRequest = false;
                    }
                    string CallerType = Convert.ToString(context.InputParameters["CallerType"]);
                    if (string.IsNullOrWhiteSpace(CallerType))
                    {
                        errorMessage.AppendLine("Caller Type is required.");
                        IsValidRequest = false;
                    }
                    else
                    {
                        if (CallerType.ToUpper() != "DEALER")
                        {
                            errorMessage.AppendLine("Please Enter The Valid Data");
                            IsValidRequest = false;
                        }
                    }
                    string DealerCode = Convert.ToString(context.InputParameters["DealerCode"]);
                    if (string.IsNullOrWhiteSpace(DealerCode))
                    {
                        errorMessage.AppendLine("Dealer Code is required.");
                        IsValidRequest = false;
                    }
                    else if (DealerCode.Length > 100)
                    {
                        errorMessage.AppendLine("Maximum length for dealer code is 100");
                        IsValidRequest = false;
                    }
                    string PreferredDate = Convert.ToString(context.InputParameters["PreferredDate"]);
                    if (!string.IsNullOrEmpty(PreferredDate))
                    {
                        if (!Regex_PreferredDate.IsMatch(PreferredDate))
                        {
                            errorMessage.AppendLine("Invalid Expected Delivery Date. Format should be [mm/dd/yyyy, mm-dd-yyyy or mm.dd.yyyy].");
                            IsValidRequest = false;
                        }
                    }

                    string ServiceAppointmentNumber = Convert.ToString(context.InputParameters["ServiceAppointmentNumber"]);
                    if (string.IsNullOrWhiteSpace(ServiceAppointmentNumber))
                    {
                        errorMessage.AppendLine("Service Appointment Number is required.");
                        IsValidRequest = false;
                    }
                    else if (ServiceAppointmentNumber.Length > 50)
                    {
                        errorMessage.AppendLine("Maximum length for Service Appointment Number is 50");
                        IsValidRequest = false;
                    }
                    QueryExpression queryExpression = new QueryExpression("msdyn_workorder");
                    queryExpression.ColumnSet = new ColumnSet("hil_purchasedfrom");
                    queryExpression.Criteria.AddCondition("hil_purchasedfrom", ConditionOperator.Equal, ServiceAppointmentNumber);
                    EntityCollection JobEntCol = service.RetrieveMultiple(queryExpression);
                    if (JobEntCol.Entities.Count > 0)
                    {
                        errorMessage.AppendLine("Service Request Already created of this Record!!");
                        IsValidRequest = false;
                    }                   
                    if (!IsValidRequest)
                    {
                        JsonResponse = JsonSerializer.Serialize(new IoTServiceCallRegistration
                        {
                            StatusCode = "204",
                            StatusDescription = errorMessage.ToString()
                        });
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                    IoTServiceCallRegistration serviceCalldata = new IoTServiceCallRegistration
                    {
                        AddressLine1 = Addressline1,
                        AddressLine2 = Addressline2,
                        AlternateNumber = AlternateNumber.Trim(),
                        CallType = CallType,
                        CallerType = CallerType,
                        CustomerFirstName = CustomerFirstName,
                        CustomerLastName = CustomerLastName,
                        CustomerMobleNo = CustomerMobleNo,
                        DealerCode = DealerCode,
                        PreferredDate = PreferredDate,//expected delivery date
                        Landmark = Landmark,
                        ChiefComplaint = ChiefComplaint,
                        Pincode = Pincode,
                        ProductSubCategoryName = ProductSubCategoryName,
                        ServiceAppointmentNumber = ServiceAppointmentNumber,
                        SourceOfJob = SourceOfJob
                    };
                    JsonResponse = JsonSerializer.Serialize(CreateServiceCallRequest(service, serviceCalldata));
                    context.OutputParameters["data"] = JsonResponse;
                }
                else
                {
                    isValid = Guid.TryParse(Convert.ToString(context.InputParameters["CustomerGuid"]), out CustomerGuid);
                    if (!isValid)
                    {
                        errorMessage.AppendLine("Invalid required customer guid.");
                        IsValidRequest = false;
                    }
                    isValid = Guid.TryParse(Convert.ToString(context.InputParameters["NOCGuid"]), out NOCGuid);
                    if (!isValid)
                    {
                        errorMessage.AppendLine("Invalid required nature of complaint.");
                        IsValidRequest = false;
                    }
                    isValid = Guid.TryParse(Convert.ToString(context.InputParameters["AddressGuid"]), out AddressGuid);
                    if (!isValid)
                    {
                        errorMessage.AppendLine("Invalid required customer address.");
                        IsValidRequest = false;
                    }
                    string PreferredDate = Convert.ToString(context.InputParameters["PreferredDate"]);
                    if (!string.IsNullOrEmpty(PreferredDate))
                    {
                        if (!Regex_PreferredDate.IsMatch(PreferredDate))
                        {
                            errorMessage.AppendLine("Invalid PreferredDate. Format should be [mm/dd/yyyy, mm-dd-yyyy or mm.dd.yyyy].");
                            IsValidRequest = false;
                        }
                    }
                    string SerialNumber = Convert.ToString(context.InputParameters["SerialNumber"]);
                    if (!string.IsNullOrWhiteSpace(SerialNumber))
                    {
                        if (SerialNumber.Length > 35)
                        {
                            errorMessage.AppendLine("Invalid Serial Number");
                            IsValidRequest = false;
                        }
                    }
                    string ProductModelNumber = Convert.ToString(context.InputParameters["ProductModelNumber"]);
                    if (!string.IsNullOrWhiteSpace(ProductModelNumber))
                    {
                        if (ProductModelNumber.Length > 100)
                        {
                            errorMessage.AppendLine("Invalid Product ModelNumber");
                            IsValidRequest = false;
                        }
                    }
                    if (!IsValidRequest)
                    {
                        JsonResponse = JsonSerializer.Serialize(new IoTServiceCallRegistration
                        {
                            StatusCode = "204",
                            StatusDescription = errorMessage.ToString()
                        });
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }

                    int PreferredPartOfDay = 0;
                    int.TryParse(Convert.ToString(context.InputParameters["PreferredPartOfDay"]), out PreferredPartOfDay);
                    string CallSubType = Convert.ToString(context.InputParameters["CallSubType"]);
                    Guid.TryParse(Convert.ToString(context.InputParameters["ProductCategoryGuid"]), out ProductCategoryGuid);
                    Guid.TryParse(Convert.ToString(context.InputParameters["ProductSubCategoryGuid"]), out ProductSubCategoryGuid);

                    IoTServiceCallRegistration serviceCalldata = new IoTServiceCallRegistration
                    {
                        CustomerMobleNo = CustomerMobleNo,
                        CustomerGuid = CustomerGuid,
                        NOCGuid = NOCGuid,
                        AddressGuid = AddressGuid,
                        SerialNumber = SerialNumber,
                        SourceOfJob = SourceOfJob,
                        ProductModelNumber = ProductModelNumber,
                        PreferredDate = PreferredDate,
                        ChiefComplaint = ChiefComplaint,
                        CallSubType = CallSubType,
                        PreferredPartOfDay = PreferredPartOfDay,
                        ProductCategoryGuid = ProductCategoryGuid,
                        ProductSubCategoryGuid = ProductSubCategoryGuid
                    };
                    JsonResponse = JsonSerializer.Serialize(IoTCreateServiceCall(service, serviceCalldata));
                    context.OutputParameters["data"] = JsonResponse;
                }
            }
            else
            {
                JsonResponse = JsonSerializer.Serialize(new IoTServiceCallRegistration
                {
                    StatusCode = "204",
                    StatusDescription = "Customer mobile number is required."
                });
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
        }

        public IoTServiceCallRegistration IoTCreateServiceCall(IOrganizationService service, IoTServiceCallRegistration serviceCalldata)
        {
            IoTServiceCallRegistration objServiceCall;
            Guid customerGuid = Guid.Empty;
            Guid callSubTypeGuid = Guid.Empty;
            Guid serviceCallGuid = Guid.Empty;
            Entity lookupObj = null;
            EntityCollection entcoll;
            QueryExpression Query;
            string customerFullName = string.Empty;
            string customerMobileNumber = string.Empty;
            string customerEmail = string.Empty;
            Guid customerAssetGuid = Guid.Empty;
            DateTime? invoiceDate = null;
            string modelName = string.Empty;
            EntityReference erProductCategory = null;
            EntityReference erProductsubcategory = null;
            EntityReference erProductsubcategorymapping = null;
            EntityReference erNatureOfComplaint = null;
            EntityReference erCustomerAsset = null;
            bool continueFlag = false;
            string fullAddress = string.Empty;
            try
            {
                if (service != null)
                {
                    if (string.IsNullOrWhiteSpace(serviceCalldata.SerialNumber))
                    {
                        if (serviceCalldata.SourceOfJob == 22)
                        {
                            if (!string.IsNullOrEmpty(serviceCalldata.ProductModelNumber))
                            {
                                Query = new QueryExpression("product");
                                Query.ColumnSet = new ColumnSet("hil_division", "hil_materialgroup", "productnumber");
                                Query.Criteria.AddCondition("productnumber", ConditionOperator.Equal, serviceCalldata.ProductModelNumber);
                                EntityCollection dataCollection = service.RetrieveMultiple(Query);
                                if (dataCollection.Entities.Count > 0)
                                {
                                    serviceCalldata.ProductCategoryGuid = dataCollection.Entities[0].GetAttributeValue<EntityReference>("hil_division").Id;
                                    serviceCalldata.ProductSubCategoryGuid = dataCollection.Entities[0].GetAttributeValue<EntityReference>("hil_materialgroup").Id;
                                }
                                else
                                {
                                    return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Invalid ProductModelNumber." };
                                }
                            }
                            else if (serviceCalldata.ProductCategoryGuid == Guid.Empty || serviceCalldata.ProductSubCategoryGuid == Guid.Empty)
                            {
                                return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Serial Number or Product Modelnumber or Product Category and Product SubCategory is required" };
                            }
                        }
                        if (serviceCalldata.ProductCategoryGuid == Guid.Empty)
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Product Category is required." };
                        }
                        else if (serviceCalldata.ProductSubCategoryGuid == Guid.Empty)
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Product Sub Category is required." };
                        }
                    }
                    Query = new QueryExpression("hil_address");
                    Query.ColumnSet = new ColumnSet("hil_addressid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_addressid", ConditionOperator.Equal, serviceCalldata.AddressGuid);
                    Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, serviceCalldata.CustomerGuid);
                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Address does not belong to Customer." };
                    }
                    Query = new QueryExpression("contact");
                    Query.ColumnSet = new ColumnSet("fullname", "emailaddress1", "mobilephone");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, serviceCalldata.CustomerMobleNo);
                    Query.Criteria.AddCondition("contactid", ConditionOperator.Equal, serviceCalldata.CustomerGuid);
                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer/Mobile No. does not exist." };
                    }
                    else
                    {
                        lookupObj = entcoll.Entities[0];
                        customerFullName = lookupObj.GetAttributeValue<string>("fullname");
                        customerEmail = lookupObj.GetAttributeValue<string>("emailaddress1");
                        customerMobileNumber = lookupObj.GetAttributeValue<string>("mobilephone"); // N
                    }
                    if (string.IsNullOrWhiteSpace(serviceCalldata.ChiefComplaint))
                    {
                        if (serviceCalldata.SourceOfJob != 22)
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer's Chief Complaint is required." };
                        }
                    }
                    //Case 1 Serial Number Exists
                    if (!string.IsNullOrWhiteSpace(serviceCalldata.SerialNumber))
                    {
                        string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_customerasset'>
                            <attribute name='msdyn_name' />
                            <attribute name='hil_customer' />
                            <attribute name='hil_productsubcategorymapping' />
                            <attribute name='hil_productcategory' />
                            <attribute name='hil_productsubcategory' />
                            <attribute name='msdyn_customerassetid' />
                            <attribute name='hil_invoicedate' />
                        <order attribute='msdyn_name' descending='false' />
	                    <filter type='and'>
                            <condition attribute='hil_customer' operator='eq' value='{" + serviceCalldata.CustomerGuid.ToString() + @"}' />
                            <condition attribute='msdyn_name' operator='eq' value='" + serviceCalldata.SerialNumber + @"' />
                        </filter>
                        </entity>
                        </fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (entcoll.Entities.Count > 0)
                        {
                            erCustomerAsset = entcoll.Entities[0].ToEntityReference();
                            modelName = entcoll.Entities[0].GetAttributeValue<string>("msdyn_name");
                            invoiceDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_invoicedate");
                            erProductCategory = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productcategory");
                            erProductsubcategory = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productsubcategory");
                            erProductsubcategorymapping = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productsubcategorymapping");
                            continueFlag = true;
                        }
                        else
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Serial number does not belong to the given customer." };
                        }
                    }
                    //Case 2 Product Category 
                    else if (serviceCalldata.ProductCategoryGuid != Guid.Empty)
                    {
                        erProductCategory = new EntityReference("product", serviceCalldata.ProductCategoryGuid);
                        erProductsubcategory = new EntityReference("product", serviceCalldata.ProductSubCategoryGuid);
                        modelName = string.Empty;
                        continueFlag = true;
                    }
                    if (!continueFlag)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Asset Serial Number/Product Category is required to proceed." };
                    }
                    int[] ArraySourceOfJob = new int[] { 12, 13, 16, 22 };
                    // Input<12> for Whatsapp, <13> for IoT Platform, <16> for Chatbot and <22> for OneWebsite.
                    if (!ArraySourceOfJob.Contains(serviceCalldata.SourceOfJob))
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Invalid source of job." };
                    }
                    #region Get Nature of Complaint
                    string fetchXML = string.Empty;

                    if (serviceCalldata.NOCGuid != Guid.Empty)
                    {
                        fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>";
                        fetchXML += "<entity name='hil_natureofcomplaint'>";
                        fetchXML += "<attribute name='hil_callsubtype' />";
                        fetchXML += "<attribute name='hil_natureofcomplaintid' />";
                        fetchXML += "<order attribute='createdon' descending='false' />";
                        fetchXML += "<filter type='and'>";
                        fetchXML += "<condition attribute='statecode' operator='eq' value='0' />";
                        fetchXML += "<condition attribute='hil_relatedproduct' operator='eq' value='{" + erProductsubcategory.Id + "}' />";
                        fetchXML += "<condition attribute='hil_natureofcomplaintid' operator='eq' value='{" + serviceCalldata.NOCGuid + "}' />";
                        fetchXML += "</filter>";
                        fetchXML += "</entity>";
                        fetchXML += "</fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXML));
                        if (entcoll.Entities.Count > 0)
                        {
                            erNatureOfComplaint = entcoll.Entities[0].ToEntityReference();
                            if (entcoll.Entities[0].Attributes.Contains("hil_callsubtype"))
                            {
                                callSubTypeGuid = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_callsubtype").Id;
                            }
                        }
                        else
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "NOC does not match with Product Sub Category." };
                        }
                    }
                    else if (serviceCalldata.SourceOfJob == 22)
                    {
                        string[] Callsubtype = new string[] { "I", "B", "D" };
                        if (Callsubtype.Contains(serviceCalldata.CallSubType))
                        {
                            fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>";
                            fetchXML += "<entity name='hil_natureofcomplaint'>";
                            fetchXML += "<attribute name='hil_callsubtype' />";
                            fetchXML += "<attribute name='hil_natureofcomplaintid' />";
                            fetchXML += "<order attribute='createdon' descending='false' />";
                            fetchXML += "<filter type='and'>";
                            fetchXML += "<condition attribute='statecode' operator='eq' value='0' />";
                            fetchXML += "<condition attribute='hil_relatedproduct' operator='eq' value='{" + erProductsubcategory.Id + "}' />";
                            fetchXML += "<condition attribute='hil_callsubtype' operator='in'>";
                            fetchXML += "<value uiname='Demo' uitype='hil_callsubtype'>{AE1B2B71-3C0B-E911-A94E-000D3AF06CD4}</value>";
                            fetchXML += "<value uiname='Installation' uitype='hil_callsubtype'>{E3129D79-3C0B-E911-A94E-000D3AF06CD4}</value></condition>";
                            fetchXML += "</filter>";
                            fetchXML += "</entity>";
                            fetchXML += "</fetch>";
                            entcoll = service.RetrieveMultiple(new FetchExpression(fetchXML));
                            if (entcoll.Entities.Count > 0)
                            {
                                if (serviceCalldata.CallSubType == "D")
                                {
                                    callSubTypeGuid = new Guid("AE1B2B71-3C0B-E911-A94E-000D3AF06CD4");//Demo
                                }
                                else
                                {
                                    callSubTypeGuid = new Guid("E3129D79-3C0B-E911-A94E-000D3AF06CD4");//Installation
                                }
                            }
                        }
                        else
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Invalid Call Sub Type." };
                        }
                    }
                    #endregion
                    #region Create Service Call
                    objServiceCall = new IoTServiceCallRegistration();
                    objServiceCall = serviceCalldata;
                    Entity enWorkorder = new Entity("msdyn_workorder");

                    if (serviceCalldata.CustomerGuid != Guid.Empty)
                    {
                        enWorkorder["hil_customerref"] = new EntityReference("contact", serviceCalldata.CustomerGuid);
                    }
                    enWorkorder["hil_customername"] = customerFullName;
                    enWorkorder["hil_mobilenumber"] = customerMobileNumber;
                    enWorkorder["hil_email"] = customerEmail;

                    if (serviceCalldata.PreferredPartOfDay > 0 && serviceCalldata.PreferredPartOfDay < 4)
                    {
                        enWorkorder["hil_preferredtime"] = new OptionSetValue(serviceCalldata.PreferredPartOfDay);
                    }
                    if (!string.IsNullOrWhiteSpace(serviceCalldata.PreferredDate))
                    {
                        string _date = serviceCalldata.PreferredDate;
                        DateTime dtInvoice = new DateTime(Convert.ToInt32(_date.Substring(6, 4)), Convert.ToInt32(_date.Substring(0, 2)), Convert.ToInt32(_date.Substring(3, 2)));
                        enWorkorder["hil_preferreddate"] = dtInvoice;
                    }
                    if (serviceCalldata.AddressGuid != Guid.Empty)
                    {
                        enWorkorder["hil_address"] = new EntityReference("hil_address", serviceCalldata.AddressGuid);
                    }

                    if (erCustomerAsset != null)
                    {
                        enWorkorder["msdyn_customerasset"] = erCustomerAsset;
                    }
                    if (modelName != string.Empty)
                    {
                        enWorkorder["hil_modelname"] = modelName;
                    }

                    if (erProductCategory != null)
                    {
                        enWorkorder["hil_productcategory"] = erProductCategory;
                    }
                    if (erProductsubcategory != null)
                    {
                        enWorkorder["hil_productsubcategory"] = erProductsubcategory;
                    }

                    Query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
                    Query.ColumnSet = new ColumnSet("hil_stagingdivisonmaterialgroupmappingid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_productcategorydivision", ConditionOperator.Equal, erProductCategory.Id);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_productsubcategorymg", ConditionOperator.Equal, erProductsubcategory.Id);
                    EntityCollection ec = service.RetrieveMultiple(Query);
                    if (ec.Entities.Count > 0)
                    {
                        enWorkorder["hil_productcatsubcatmapping"] = ec.Entities[0].ToEntityReference();
                    }

                    EntityCollection entCol;
                    Query = new QueryExpression("hil_consumertype");
                    Query.ColumnSet = new ColumnSet("hil_consumertypeid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "B2C");
                    entCol = service.RetrieveMultiple(Query);
                    if (entCol.Entities.Count > 0)
                    {
                        enWorkorder["hil_consumertype"] = entCol.Entities[0].ToEntityReference();
                    }

                    Query = new QueryExpression("hil_consumercategory");
                    Query.ColumnSet = new ColumnSet("hil_consumercategoryid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "End User");
                    entCol = service.RetrieveMultiple(Query);
                    if (entCol.Entities.Count > 0)
                    {
                        enWorkorder["hil_consumercategory"] = entCol.Entities[0].ToEntityReference();
                    }

                    if (erNatureOfComplaint != null)
                    {
                        enWorkorder["hil_natureofcomplaint"] = erNatureOfComplaint;
                    }
                    if (callSubTypeGuid != Guid.Empty)
                    {
                        enWorkorder["hil_callsubtype"] = new EntityReference("hil_callsubtype", callSubTypeGuid);
                    }
                    enWorkorder["hil_quantity"] = 1;
                    enWorkorder["hil_customercomplaintdescription"] = serviceCalldata.ChiefComplaint;
                    enWorkorder["hil_callertype"] = new OptionSetValue(910590001);// {SourceofJob:"Customer"}

                    enWorkorder["hil_sourceofjob"] = new OptionSetValue(serviceCalldata.SourceOfJob); // SourceofJob:[{"12": "WhatsApp"} ,{"13","IoT Platform"},{"16","Chatbot"}]

                    enWorkorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {ServiceAccount:"Dummy Account"}
                    enWorkorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {BillingAccount:"Dummy Account"}
                    //enWorkorder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid("0F5E8009-3BFD-E811-A94C-000D3AF0694E")); // {Primary Incident Type:"Installation&nbsp;-Decorative&nbsp;FAN&nbsp;CF"}

                    serviceCallGuid = service.Create(enWorkorder);
                    if (serviceCallGuid != Guid.Empty)
                    {
                        objServiceCall.JobGuid = serviceCallGuid;
                        objServiceCall.JobId = service.Retrieve("msdyn_workorder", serviceCallGuid, new ColumnSet("msdyn_name")).GetAttributeValue<string>("msdyn_name");
                        objServiceCall.StatusCode = "200";
                        objServiceCall.StatusDescription = "OK";
                    }
                    else
                    {
                        objServiceCall.StatusCode = "204";
                        objServiceCall.StatusDescription = "FAILURE !!! Something went wrong";
                    }
                    return objServiceCall;
                    #endregion
                }
                else
                {
                    objServiceCall = new IoTServiceCallRegistration { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    return objServiceCall;
                }
            }
            catch (Exception ex)
            {
                objServiceCall = new IoTServiceCallRegistration { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                return objServiceCall;
            }
        }
        public IoTServiceCallRegistration CreateServiceCallRequest(IOrganizationService service, IoTServiceCallRegistration serviceCalldata)
        {
            try
            {
                Guid businessmappingId = Guid.Empty;
                QueryExpression query = new QueryExpression("hil_pincode");
                query.ColumnSet = new ColumnSet(false);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, serviceCalldata.Pincode);
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                EntityCollection entcollpincode = service.RetrieveMultiple(query);
                if (entcollpincode.Entities.Count > 0)
                {
                    query = new QueryExpression("hil_businessmapping");
                    query.TopCount = 1;
                    query.ColumnSet.AddColumns("hil_businessmappingid", "hil_pincode");
                    query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, entcollpincode.Entities[0].Id);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    EntityCollection businessmapping = service.RetrieveMultiple(query);
                    if (businessmapping.Entities.Count > 0)
                    {
                        businessmappingId = businessmapping.Entities[0].Id;
                    }
                    else
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "The pincode mapping is absent; please reach out to the system administrator for assistance." };
                    }
                }
                else
                {
                    return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Not a valid Pincode." };
                }
                #region Create Service Call
                Entity enWorkorder = new Entity("msdyn_workorder");

                #region Customer_Creation
                Guid contactId = Guid.Empty;
                QueryExpression Query = new QueryExpression("contact");
                Query.ColumnSet = new ColumnSet("fullname", "emailaddress1", "mobilephone");
                Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, serviceCalldata.CustomerMobleNo);
                EntityCollection entcoll = service.RetrieveMultiple(Query);
                if (entcoll.Entities.Count > 0)
                {
                    contactId = entcoll.Entities[0].Id;
                    enWorkorder["hil_customername"] = entcoll.Entities[0].Contains("fullname") ? entcoll.Entities[0].GetAttributeValue<string>("fullname") : "";
                    enWorkorder["hil_mobilenumber"] = entcoll.Entities[0].Contains("mobilephone") ? entcoll.Entities[0].GetAttributeValue<string>("mobilephone") : "";
                    enWorkorder["hil_email"] = entcoll.Entities[0].Contains("emailaddress1") ? entcoll.Entities[0].GetAttributeValue<string>("emailaddress1") : "";
                }
                else
                {
                    Entity entConsumer = new Entity("contact");
                    entConsumer["mobilephone"] = serviceCalldata.CustomerMobleNo;
                    entConsumer["firstname"] = serviceCalldata.CustomerFirstName;
                    entConsumer["lastname"] = serviceCalldata.CustomerLastName ?? "";
                    entConsumer["address1_telephone3"] = serviceCalldata.AlternateNumber ?? "";
                    entConsumer["hil_consumersource"] = new OptionSetValue(21); //MFR
                    entConsumer["hil_subscribeformessagingservice"] = true;
                    contactId = service.Create(entConsumer);
                }
                #endregion

                if (contactId != Guid.Empty)
                {
                    enWorkorder["hil_customerref"] = new EntityReference("contact", contactId);
                    enWorkorder["hil_mobilenumber"] = serviceCalldata.CustomerMobleNo;
                    enWorkorder["hil_alternate"] = serviceCalldata.AlternateNumber;
                }
                if (!string.IsNullOrWhiteSpace(serviceCalldata.PreferredDate))
                {
                    string _date = serviceCalldata.PreferredDate;
                    DateTime preferreddate = new DateTime(Convert.ToInt32(_date.Substring(6, 4)), Convert.ToInt32(_date.Substring(0, 2)), Convert.ToInt32(_date.Substring(3, 2)));
                    enWorkorder["hil_preferreddate"] = preferreddate;
                }

                #region Address_Creation
                Entity address = new Entity("hil_address");
                address["hil_customer"] = new EntityReference("contact", contactId);
                address["hil_street1"] = serviceCalldata.AddressLine1;
                address["hil_street2"] = serviceCalldata.AddressLine2;
                address["hil_street3"] = serviceCalldata.Landmark;
                address["hil_addresstype"] = new OptionSetValue(1);
                address["hil_businessgeo"] = new EntityReference("hil_businessmapping", businessmappingId);
                Guid Addressid = service.Create(address);
                #endregion

                if (Addressid != Guid.Empty)
                {
                    enWorkorder["hil_address"] = new EntityReference("hil_address", Addressid);
                }
                Query = new QueryExpression("product");
                Query.ColumnSet = new ColumnSet("hil_division");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("name", ConditionOperator.Equal, serviceCalldata.ProductSubCategoryName);
                Query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 3);
                EntityCollection entCol = service.RetrieveMultiple(Query);
                if (entCol.Entities.Count > 0)
                {
                    enWorkorder["hil_productsubcategory"] = new EntityReference("product", entCol.Entities[0].Id);

                    Query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
                    Query.ColumnSet = new ColumnSet("hil_stagingdivisonmaterialgroupmappingid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_productcategorydivision", ConditionOperator.Equal, entCol.Entities[0].GetAttributeValue<EntityReference>("hil_division").Id);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_productsubcategorymg", ConditionOperator.Equal, entCol.Entities[0].Id);
                    EntityCollection ec = service.RetrieveMultiple(Query);
                    if (ec.Entities.Count > 0)
                    {
                        enWorkorder["hil_productcatsubcatmapping"] = ec.Entities[0].ToEntityReference();
                    }
                    string _fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='hil_natureofcomplaint'>
                                        <attribute name='hil_name' />
                                        <attribute name='hil_natureofcomplaintid' />
                                        <attribute name='hil_callsubtype' />
                                        <order attribute='hil_name' descending='false' />
                                        <filter type='and'>
                                            <condition attribute='hil_relatedproduct' operator='eq' value='{entCol.Entities[0].Id}' />
                                            <condition attribute='statecode' operator='eq' value='0' />
                                        </filter>
                                        <link-entity name='hil_callsubtype' from='hil_callsubtypeid' to='hil_callsubtype' link-type='inner' alias='ad'>
                                        <filter type='and'>
                                            <condition attribute='hil_name' operator='eq' value='{serviceCalldata.CallType}' />
                                        </filter>
                                        </link-entity>
                                        </entity>
                                        </fetch>";

                    EntityCollection _natureofcomplaintColl = service.RetrieveMultiple(new FetchExpression(_fetchQuery));
                    if (_natureofcomplaintColl.Entities.Count > 0)
                    {
                        enWorkorder["hil_natureofcomplaint"] = _natureofcomplaintColl.Entities[0].ToEntityReference();
                        enWorkorder["hil_callsubtype"] = _natureofcomplaintColl.Entities[0].GetAttributeValue<EntityReference>("hil_callsubtype");
                    }
                }
                enWorkorder["hil_consumertype"] = new EntityReference("hil_consumertype", new Guid("484897de-2abd-e911-a957-000d3af0677f")); //B2C
                enWorkorder["hil_consumercategory"] = new EntityReference("hil_consumercategory", new Guid("baa52f5e-2bbd-e911-a957-000d3af0677f")); //End User
                enWorkorder["hil_quantity"] = 1;
                enWorkorder["hil_customercomplaintdescription"] = serviceCalldata.ChiefComplaint;
                enWorkorder["hil_callertype"] = new OptionSetValue(910590000);//  Dealer
                enWorkorder["hil_newserialnumber"] = serviceCalldata.DealerCode;
                enWorkorder["hil_sourceofjob"] = new OptionSetValue(serviceCalldata.SourceOfJob); //MFR and Flipkart
                enWorkorder["hil_preferredtime"] = new OptionSetValue(1);//MFR set default Prefered day [Morning]
                enWorkorder["hil_purchasedfrom"] = serviceCalldata.ServiceAppointmentNumber;
                enWorkorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {ServiceAccount:"Dummy Account"}
                enWorkorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {BillingAccount:"Dummy Account"}

                Guid serviceCallGuid = service.Create(enWorkorder);
                if (serviceCallGuid != Guid.Empty)
                {
                    return new IoTServiceCallRegistration
                    {
                        StatusCode = "200",
                        JobGuid = serviceCallGuid,
                        JobId = service.Retrieve("msdyn_workorder", serviceCallGuid, new ColumnSet("msdyn_name")).GetAttributeValue<string>("msdyn_name"),
                        StatusDescription = "Service call request registered successfully."
                    };
                }
                else
                {
                    return new IoTServiceCallRegistration
                    {
                        StatusCode = "204",
                        StatusDescription = "Something went wrong; please reach out to the system administrator for assistance."
                    };
                }
                #endregion
            }
            catch (Exception ex)
            {
                return new IoTServiceCallRegistration
                {
                    StatusCode = "200",
                    StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper()
                };
            }
        }
    }
    public class IoTServiceCallRegistration
    {
        public string SerialNumber { get; set; }
        public string ProductModelNumber { get; set; }
        public Guid NOCGuid { get; set; }
        public string NOCName { get; set; }
        public Guid ProductCategoryGuid { get; set; }
        public Guid ProductSubCategoryGuid { get; set; }
        public string ChiefComplaint { get; set; }
        public Guid AddressGuid { get; set; }
        public Guid AssetGuid { get; set; }
        public string CustomerMobleNo { get; set; }
        public Guid CustomerGuid { get; set; }
        public Guid JobGuid { get; set; }
        public string JobId { get; set; }
        //public string ImageBase64String { get; set; }
        //public int ImageType { get; set; }
        public int SourceOfJob { get; set; }
        public string PreferredDate { get; set; }
        public int PreferredPartOfDay { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
        public string DealerCode { get; set; }
        public string DealerName { get; set; }
        public string CustomerName { get; set; }
        public string AddressLine1 { get; set; }
        public string Pincode { get; set; }
        public string PreferredLanguage { get; set; }
        public string CallSubType { get; set; }

        // New parameter added for Croma

        public string AddressLine2 { get; set; }
        public string Landmark { get; set; }
        public string AlternateNumber { get; set; }
        public string CustomerFirstName { get; set; }
        public string CustomerLastName { get; set; }
        public string ProductSubCategoryName { get; set; }
        public string CallerType { get; set; } // Dealer
        public string CallType { get; set; } // e.g. Instollation, Breakdown
        public string ServiceAppointmentNumber { get; set; }
    }

}