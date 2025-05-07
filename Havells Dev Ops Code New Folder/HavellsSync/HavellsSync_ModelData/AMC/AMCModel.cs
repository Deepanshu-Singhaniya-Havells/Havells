using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.Product;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_ModelData.AMC
{

    #region Common Models
    public class ConsumerProfile
    {
        public string mobileNumber { get; set; }
        public string fullName { get; set; }
        public string emailId { get; set; }
        public string state { get; set; }
        public string pinCode { get; set; }
        public string salesOffice { get; set; }
        public string branchMemorandumCode { get; set; }
        public string fullAddress { get; set; }
    }
    public class ResInvoiceInfo
    {
        public string TransactionID { get; set; }
        public Guid InvoiceID { get; set; }
        public string productinfo { get; set; }
        public string HashCode { get; set; }
        public string key { get; set; }
        public string salt { get; set; }
        public bool ResultStatus { get; set; }
        public string ResultMessage { get; set; }

    }
    public class SourceTypeParam
    {
        public string SourceType { get; set; }
    }
    public class TokenExpires
    {
        public int StatusCode { get; set; }
    }
    public class IntegrationConfiguration
    {
        public string url { get; set; }
        public string userName { get; set; }
        public string password { get; set; }
    }

    #endregion

    #region Get Warranty Content
    public class WarrantyContentRes : TokenExpires
    {
        //public string SourceType { get; set; }
        public List<AMCCatagory> AMCCatagories { get; set; }
        public List<WarrantyMediaContent> WarrantyMediaContents { get; set; }
        public List<WarrantyDiscountBanner> WarrantyDiscountBanners { get; set; }
    }
    public class AMCCatagory
    {
        public string Icon { get; set; }
        public string CategoryName { get; set; }
        public Guid CategoryId { get; set; }
    }
    public class WarrantyMediaContent
    {
        public string Icon { get; set; }
        public string Content { get; set; }
    }
    public class WarrantyDiscountBanner
    {
        public int Index { get; set; }
        public string URL { get; set; }
    }
    #endregion

    #region Get AMC Plan
    public class AMCPlanParam
    {
        public string ModelNumber { get; set; }
        public string SourceType { get; set; }
        public string AddressId { get; set; }
        public string CustomerAssestId { get; set; }
    }
    public class AMCPlanInfo
    {
        public Guid PlanId { get; set; }
        public string PlanName { get; set; }
        public string PlanPeriod { get; set; }
        public string MRP { get; set; }
        public string DiscountPercent { get; set; }
        public string EffectivePrice
        {
            get
            {
                if (DiscountPercent != null)
                {
                    return decimal.Round(Convert.ToDecimal(MRP) - ((Convert.ToDecimal(MRP) * Convert.ToDecimal(DiscountPercent)) / 100), 2).ToString();
                }
                else
                {
                    return decimal.Round(Convert.ToDecimal(MRP), 2).ToString();
                }
            }
        }
        public string Coverage { get; set; }
        public string NonCoverage { get; set; }
        public string PlanTCLink { get; set; }
    }
    public class AMCPlanRes : TokenExpires
    {
        public string ModelNumber { get; set; }
        public List<AMCPlanInfo> AMCPlanInfo { get; set; }
    }
    #endregion

    #region Get Payment Status
    public class PaymentStatusParam
    {
        public Guid InvoiceID { get; set; }
        public string SourceType { get; set; }
    }

    public class PaymentStatusRes : TokenExpires
    {
        public string PaymentStatus { get; set; }
    }

    public class StatusRequest
    {
        public string PROJECT { get; set; }
        public string command { get; set; }
        public string var1 { get; set; }
    }
    public class StatusResponse
    {
        public string status { get; set; }
        public string msg { get; set; }
        public List<TransactionDetail> transaction_details { get; set; }
    }
    public class TransactionDetail
    {
        public string mihpayid { get; set; }
        public string request_id { get; set; }
        public string bank_ref_num { get; set; }
        public string amt { get; set; }
        public string transaction_amount { get; set; }
        public string txnid { get; set; }
        public string additional_charges { get; set; }
        public string productinfo { get; set; }
        public string firstname { get; set; }
        public string bankcode { get; set; }
        public string udf1 { get; set; }
        public string udf3 { get; set; }
        public string udf4 { get; set; }
        public string udf5 { get; set; }
        public string field2 { get; set; }
        public string field9 { get; set; }
        public string error_code { get; set; }
        public string addedon { get; set; }
        public string payment_source { get; set; }
        public string card_type { get; set; }
        public string error_Message { get; set; }
        public string net_amount_debit { get; set; }
        public string disc { get; set; }
        public string mode { get; set; }
        public string PG_TYPE { get; set; }
        public string card_no { get; set; }
        public string udf2 { get; set; }
        public string status { get; set; }
        public string unmappedstatus { get; set; }
        public string Merchant_UTR { get; set; }
        public string Settled_At { get; set; }
    }
    #endregion

    #region getRegistredProductList
    public class AMCRegisterdProduct
    {
        public string SerialNumber { get; set; }
        public string InvoiceDate { get; set; }
        public string InvoiceNumber { get; set; }
        public decimal InvoiceValue { get; set; }
        public string ProductCategory { get; set; }
        public Guid ProductCategoryId { get; set; }
        public string ModelCode { get; set; }
        public string ModelId { get; set; }
        public string ModelName { get; set; }
        public string ProductSubCategory { get; set; }
        public Guid ProductSubCategoryId { get; set; }
        public List<AMCProductWarranty> ProductWarranty { get; set; }
        public Guid ProductGuid { get; set; }
        public string WarrantyEndDate { get; set; }
        public string WarrantyStatus { get; set; }
        public string WarrantySubStatus { get; set; }
    }

    public class AMCProductDeatilsList : TokenExpires
    {
        public List<AMCRegisterdProduct> ProductList { get; set; }
    }
    public class AMCProductWarranty
    {
        public string WarrantyEndDate { get; set; }
        public string WarrantySpecifications { get; set; }
        public string WarrantyStartDate { get; set; }
        public string WarrantyType { get; set; }
    }
    #endregion

    #region Get AMC Orders

    public class AMCOrdersParam
    {
        public string CustomerGuId { get; set; }
        public string SourceType { get; set; }
    }
    public class AMCOrdersListRes : TokenExpires
    {
        public List<AMCOrders> AMCOrders { get; set; }
    }
    public class AMCOrders
    {
        public string InvoiceId { get; set; }
        public string InvoiceDate { get; set; }
        public string InvoiceValue { get; set; }
        public string PaymentStatus { get; set; }
        public string InvoiceDescription { get; set; }
    }
    public class TranscationDetails
    {
        public Guid InvoiceId { get; set; }
        public List<TranscationHistory> Transcation { get; set; }
    }
    public class TranscationHistory
    {
        public string InvoiceId { get; set; }
        public string Transactionid { get; set; }
        public string PlanName { get; set; }
        public string ProductName { get; set; }
        public string Amount { get; set; }
        public string PlanDuration { get; set; }
        public string PaymentStatus { get; set; }
        public string TransactionDate { get; set; }
        public string InfoMessage { get; set; }
    }
    public class TranscationHistoryWithOrder
    {
        public string InvoiceId { get; set; }
        public string Transactionid { get; set; }
        public string PlanName { get; set; }
        public string ProductName { get; set; }
        public string Amount { get; set; }
        public string PlanDuration { get; set; }
        public string PaymentStatus { get; set; }
        public string TransactionDate { get; set; }
        public string InfoMessage { get; set; }
        public Nullable<DateTime> TransactionOrder { get; set; }
    }


    #endregion

    #region Initiate Payment
    public class InitiatePaymentParam
    {
        public string AMCPlanID { get; set; }
        public string AssestId { get; set; }
        public string DOP { get; set; }
        public decimal Price { get; set; }
        public string AddressID { get; set; }
        public string SourceType { get; set; }
    }
    public class InitiatePaymentRes : TokenExpires
    {
        public string TransactionID { get; set; }
        public string InvoiceID { get; set; }
        public string HashName { get; set; }
        public string Key { get; set; }
        public string Salt { get; set; }
        public string MobileNumber { get; set; }
        public string MamorandumCode { get; set; }
        public string Emailid { get; set; }
        public string UserName { get; set; }
        public string Surl { get; set; }
        public string Furl { get; set; }
    }

    public class PaymentLinkReponse
    {
        public string TransactionId { get; set; }
        public string productinfo { get; set; }
        public string HashCode { get; set; }
        public string key { get; set; }
        public string salt { get; set; }
        public string url { get; set; }
        public bool ResultStatus { get; set; }
        public string ResultMessage { get; set; }
    }
    public class RemotePaymentLinkDetails
    {

        public string amount { get; set; }

        public string txnid { get; set; }

        public string productinfo { get; set; }

        public string firstname { get; set; }

        public string email { get; set; }

        public string phone { get; set; }

        public string address1 { get; set; }

        public string city { get; set; }

        public string state { get; set; }

        public string country { get; set; }

        public string zipcode { get; set; }

        public string template_id { get; set; }

        public string validation_period { get; set; }

        public string send_email_now { get; set; }

        public string send_sms { get; set; }

        public string time_unit { get; set; }
    }
    public class SendPayNowRequest
    {
        public string businessType { get; set; }
        public string paymentgateway_type { get; set; }
        public string IM_PROJECT { get; set; }
        public RequestData RequestData { get; set; }
    }
    public class RequestData
    {
        public string txnid { get; set; }
        public string amount { get; set; }
        public string productinfo { get; set; }
        public string firstname { get; set; }
        public string email { get; set; }
        public string udf1 { get; set; }
        public string udf2 { get; set; }
        public string udf3 { get; set; }
        public string udf4 { get; set; }
        public string udf5 { get; set; }

    }
    public class PayNowResponse
    {
        public string HashCode { get; set; }
        public string key { get; set; }
        public string salt { get; set; }
        public string status { get; set; }
        public string message { get; set; }

    }
    public class SendPaymentUrlRequest
    {
        public string PROJECT { get; set; }
        public string command { get; set; }
        public RemotePaymentLinkDetails RemotePaymentLinkDetails { get; set; }
    }
    public class SendPaymentUrlResponse
    {
        public string Email_Id { get; set; }
        public string Transaction_Id { get; set; }
        public string URL { get; set; }
        public string Status { get; set; }
        public string Phone { get; set; }
        public string StatusCode { get; set; }
        public string msg { get; set; }
    }
    #endregion
}
