using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_ModelData.Common
{
    public static class CommonMessage
    {
        public static string PaymentlinkalreadysentMsg { get; } = "Payment url has been already sent.";
        public static string InvoiceNotCreatedMsg { get; } = "Invoice not created.";
        public static string PaymenyreciveMsg { get; } = "Payment is already received against given details.";
        public static string Havells_PluginMsg { get; } = "Havells_Plugin.Helper.GetGuidbyName.";
        public static string InvalidFileType { get; } = "Invalid file type.";
        public static string InvalidMobileNumber { get; } = "Invalid mobile number";
        public static string InvoiceNoNotFound { get; } = "Order not found.";
        public static string InvalidInvoiceNumber { get; } = "Invalid invoice number.";
        public static string InvalidinvoiceBase64 { get; } = "Invalid invoice Base64.";
        public static string InvalidProductGuidMsg { get; } = "Invalid product guid.";
        public static string InvalidCustomerGuid { get; } = "Invalid customer guid.";
        public static string UnauthorizationMsg { get; } = "Unauthorization to access!!!";
        public static string InvalidSourcetypeMsg { get; } = "Invalid source type.";
        public static string SourcetypeMsg { get; } = "Source type is required.";
        public static string LoginUserIdMsg { get; } = "Login user id is required.";
        public static string SuccessMsg { get; } = "Success";
        public static string AccessdeniedMsg { get; } = "Access denied!!! ";
        public static string UserdoesnotexistMsg { get; } = "User does not exist.";
        public static string BadRequestMsg { get; } = "Invalid request!!!";
        public static string ServiceUnavailableMsg { get; } = "D365 service unavailable.";
        public static string NoRecordFound { get; } = "No product found against the user.";
        public static string ErrorattachingnotesMsg { get; } = "Error attaching notes ";
        public static string InternalServerErrorMsg { get; } = "D365 internal server error : ";
        public static string ExtSourceTypeMsg { get; } = "API is not extended to source type.";
        public static string InvalidsessionMsg { get; } = "Invalid access token.";
        public static string SessionexpiredMsg { get; } = "Session has been expired";
        public static string SerialnumberMsg { get; } = "Asset serial number is required.";
        public static string InvalidSerialnumMsg { get; } = "Invalid serial number (IDU).";
        public static string SAPapiConfigMsg { get; } = "SAP api Config not found.";
        public static string Serialnumberalreadyexist { get; } = "Provided serial number (IDU) already exists.";
        public static string Modelnotexist { get; } = "Model does't exist in  D365";
        public static string InvalidModelNumber { get; } = "Invalid model number.";
        public static string ModelNumberRequired { get; } = "ModelNumber Required";
        public static string MobileNumberMsg { get; } = "Mobile number not found";
        public static string Division_materialgroupMsg { get; } = "Division or material group mapping is missing.";
        public static string ProductGuidMsg { get; } = "Product guid is required.";
        public static string CustomerguidMsg { get; } = "Customer guid is required.";
        public static string InvoicedateMsg { get; } = "Invoice date(yyyy-MM-dd) is required.";
        public static string InvalidInvoiceDateMsg { get; } = "Invalid invoice date. Please try with this format(yyyy-MM-dd).";
        public static string FiletypeMsg { get; } = "File type(image/jpeg[0]|image/png[1]|application/pdf[2]) is required.";
        public static string CustomerNotExitMsg { get; } = "Customer does not exist.";
        public static string FotFoundMsg { get; } = "Requested resource does not exist";

        public static string MandatotyAssestIdMsg { get; } = "Assest id is required";
        public static string InvalidAssestIdMsg { get; } = "Invalid Customer Assit id";
        public static string MandatotyPriceMsg { get; } = "Price is required and should be greater than 0";
        public static string MandatotySourseTypeMsg { get; } = "Sourse type is required and must be 6";
        public static string MandatotyAddressMsg { get; } = "Address is required";
        public static string MandatotyAddressIdMsg { get; } = "Address guid required";
        public static string MandatotyDateMsg { get; } = "Invalid date format. It should be (yyyy-MM-dd)";
        public static string PaymentPriceisnotvalid { get; } = "Invalid price amount for payment";
        public static string InvalidPaymentDateMsg { get; } = "Invalid paln id";
        public static string PendingMsg { get; } = "InCase of failure the amount will be deposited back to your account in 24 hours";
        public static string FailedMsg { get; } = "Refund initiated: the amount will be reflected in your account in 24 hours";
        #region EPOS
        public static string ServiceCallLineItemrequired { get; } = "Service Call Line Item is required";
        public static string Consent { get; } = "Consent is required";
        public static string Address { get; } = "Address line 1 required";
        public static string PinCode { get; } = "Pin Code required";
        public static string SerialandSkuCode { get; } = "Serial number or Sku code required";

        public static string PreferredDateofService { get; } = "Preferred Date of Service required";
        public static string PreferredTimeofService { get; } = "Preferred Time of Service required";
        public static string Installationrequired { get; } = "Installation required";
        public static string JobIdNotFound { get; } = "Given Job Id not Found";
        #endregion

        #region Consumer App rating
        public static string RatingMsg { get; } = "Invalid Rating!!";
        public static string InvalidSourceTypeMsg { get; } = "Invalid Source Type!!";
        #endregion

        public static string PreferredDaytime { get; } = "Please enter valid preferred Daytime";
        public static string PreferredDate { get; } = "Please enter valid preferred Date";
        public static string OrderValue { get; } = "Please enter order Value";
        public static string ReceiptAmount { get; } = "Please enter receipt amount";
        public static string PaymentType { get; } = "Please select payment type";
        public static string ServiceList { get; } = "Please add at list one order in order line";
        public static string AddressNotExitMsg { get; } = "Address does not exist";
        public static string ServiceNotExits { get; } = "Service does not exist";
        public static string QuantityValid { get; } = "Please enter quantity";
        public static string MRPValid { get; } = "Please enter MRP value";
        public static string OrderNotCreated { get; } = "Order is not created";
        public static string InvalidOrderMsg { get; } = "Invalid order id";
        public static string Invalidaddress { get; } = "Invalid address guid";
        public static string ServiceNameRequird { get; } = "Service name required";
        public static string InvalidDiscount { get; } = "Invalid discount";
        public static string InvalidDiscountMsg { get; } = "Discount amount should not be greater than Receipt amount.";
        public static string InvalidReceiptorOrderAmount { get; } = "Receipt amount should not be greater than order amount.";
        public static string InvalidCategory { get; } = "Invalid Category";
        public static string InvalidSubCategory { get; } = "Invalid Sub-Category";
        public static string NoRecordExit { get; } = "No record found.";
        public static string DiscountAmount_PercentageMsg { get; } = "please provide discount amount or discount percentage.";
        public static string Requiremessage { get; } = "Required all fields";
        public static string InvalidAMCPlanID { get; } = "Invalid AMC plan guid";
        public static string MandatoryAMCPlanID { get; } = "AMC plan id required";
        public static string PriceValidationMessage { get; } = "Invalid request, please try again later.";

        #region MFRServiceJobs Messages
        public static string DealerCodeRequired { get; } = "Dealer code is required.";
        public static string DealerCodeLength { get; } = "Dealer code length should be between 5 to 10.";
        public static string FromdateRequired { get; } = "From date(Format:yyyy-MM-dd) is required.";
        public static string TodateRequired { get; } = "To date(Format:yyyy-MM-dd) is required.";
        public static string JobRequired { get; } = "Job Id is required.";

        #endregion

    }
}
