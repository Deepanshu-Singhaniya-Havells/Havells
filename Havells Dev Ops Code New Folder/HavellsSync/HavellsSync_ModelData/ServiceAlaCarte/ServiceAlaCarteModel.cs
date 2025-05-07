using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Web.UI.WebControls;
using System.Xml.Linq;

namespace HavellsSync_ModelData.ServiceAlaCarte
{

    public class ProductCatagories
    {
        public string CategoryImageUrl { get; set; }
        public string ProductCategory { get; set; }
        public string ProductCategoryId { get; set; }
        public string ProductSubCategory { get; set; }
        public string ProductSubCategoryId { get; set; }

    }

    public class param<T>
    {
        public T data { get; set; }
    }
    public class SourceTypes
    {
        public string SourceType { get; set; }

    }

    public class ServiceListParam
    {
        public string SourceType { get; set; }
        public string ProuctCategoryId { get; set; }
        public string ProductSubCategoryId { get; set; }

    }

    public class PaymentRetryParam
    {
        public string SourceType { get; set; }
        public string OrderId { get; set; }
        public string PaymentType { get; set; }
    }
    public class OrderDetailRequest
    {
        public string SourceType { get; set; }
        public string OrderGuid { get; set; }
    }
    public class ServiceOrderDetails
    {
        public ServiceOrder serviceOrder { get; set; }
        public string CustomerName { get; set; }
    }

    public class ServiceOrder
    {
        public string OrderId { get; set; }
        public string PreferredDate { get; set; }
        public string PreferredDateTime { get; set; }
        public string AddressType { get; set; }
        public string ProductCategory { get; set; }
        public string ProductCategoryId { get; set; }
        public string CategoryImageUrl { get; set; }
        public List<BookServies> BookedServiceList { get; set; }
        public PaymentDetails PaymentInfo { get; set; }
    }

    public class BookServies
    {
        public string ServiceName { get; set; }

        public string ServiceId { get; set; }

        public double MRP { get; set; }

        public int Quantity { get; set; }
    }

    public class PaymentDetails
    {
        public string TransactionID { get; set; }
        public string HashName { get; set; }
        public string Key { get; set; }
        public string Amount { get; set; }
        public string Salt { get; set; }
        public string MamorandumCode { get; set; }
        public string MobileNumber { get; set; }
        public string Emailid { get; set; }
        public string CustomerName { get; set; }
        public string Surl { get; set; }
        public string Furl { get; set; }
        //public string udf1 { get; set; }
        //public string udf2 { get; set; }
    }

    public class PayNowRequest
    {
        public string businessType { get; set; }
        public string paymentgateway_type { get; set; }
        public string IM_PROJECT { get; set; }
        public PaymentDetails RequestData { get; set; }
    }
    public class PayNowPayu
    {
        public string HashCode { get; set; }
        public string key { get; set; }
        public string salt { get; set; }
        public string status { get; set; }
        public string message { get; set; }

    }

    public class ServiceAlaCartePlanList
    {
        public Guid SubCategoryId { get; set; }
        public string SubCategoryName { get; set; }
        public string ServiceId { get; set; }
        public string ServiceName { get; set; }
        public string ServiceDuration { get; set; }
        public string MRP { get; set; }
        public string DiscountPercent { get; set; }
        public string DiscountValue
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(DiscountPercent))
                {
                    return (decimal.Round((Convert.ToDecimal(MRP) * Convert.ToDecimal(DiscountPercent) / 100), 2)).ToString();
                }
                else
                {
                    return "0.00";
                }
            }
        }


        public string Included { get; set; }
        public string Excluded { get; set; }

    }

    public class ServiceAlaCartePlanInfo
    {
        public List<ServiceAlaCartePlanList> ServiceAlaCartePlanList { get; set; }
        public string ServiceBannerUrl { get; set; }

    }

    public class CreateOrder
    {
        public string CustomerId { get; set; }
        public string AddressId { get; set; }
        public string PreferredDate { get; set; }
        public string PreferredDateTime { get; set; }
        public string OrderValue { get; set; }
        public string DiscountAmount { get; set; }
        public string ReceiptAmount { get; set; }
        public string SourceType { get; set; }
        public string PaymentType { get; set; }
        public List<Services> ServiceList { get; set; }
    }
    public class BookedService
    {
        public string ProductCategory { get; set; }
        public string ProductCategoryId { get; set; }
        public string ServiceName { get; set; }
        public string ServiceId { get; set; }
        public string ServicePrice { get; set; }
        public string CategoryImageUrl { get; set; }
        public int ServiceDisplayIndex { get; set; }
        public string ProductSubCategory { get; set; }
        public string ProductSubCategoryId { get; set; }

    }
    public class Services
    {
        public string ServiceName { get; set; }
        public string ServiceId { get; set; }
        public string Quantity { get; set; }
        public string MRP { get; set; }
        public string DiscountPercent { get; set; }
        public string DiscountValue { get; set; }
    }
    public class ReschuduleService
    {
        public string orderId { get; set; }
        public string PreferredDate { get; set; }
        public string PreferredDateTime { get; set; }
        public string SourceType { get; set; }
    }
    //public class Source
    //{
    //    public string SourceType { get; set; }
    //}
    public class ServiceRequest
    {
        public string CustomerId { get; set; }
        public string SourceType { get; set; }
    }
    public class OrdersList
    {
        public string ProductCategory { get; set; }
        public string OrderGUID { get; set; }
        public string NumberOfServices { get; set; }
        public string OrderAmount { get; set; }
        public string OrderStatus { get; set; }
        public string AssignTo { get; set; }
        public string EngMobileNo { get; set; }
        public string OrderId { get; set; }
    }
    public class ServiceRequestData
    {
        public string ProductCategory { get; set; }
        public string ServiceLocation { get; set; }
        public string ServiceDateTime { get; set; }
        public string OrderId { get; set; }
        public string OrderAmount { get; set; }
        public string PaymentStatus { get; set; }
        public string PaymentMode { get; set; }
        public string DownloadInvoiceLink { get; set; }
        public List<ServiceJobDetail> ServiceJobDetails { get; set; }
    }

    public class ServiceJobDetail
    {
        public string ServiceName { get; set; }
        public string JobStatus { get; set; }
        public string AssignTo { get; set; }
        public string EngMobileNo { get; set; }
        public double ServiceAmount { get; set; }
        public string JobId { get; set; }
        public List<JobStatusTracker> JobStatusTracker { get; set; }
    }
    public class JobStatusTracker
    {
        public string JobStatus { get; set; }
        public string JobStatusDateTime { get; set; }
    }
    public class EntityInfo
    {
        public Guid Id { get; set; }
        public string Name { get; set; }

    }
    public class LTTABLE
    {
        public string MATNR { get; set; }
        public string KSCHL { get; set; }
        public string KBETR { get; set; }
        public string KONWA { get; set; }
        public string DATAB { get; set; }
        public string DATBI { get; set; }
    }
    public class LTTABLEResponse
    {
        public string MATNR { get; set; }
        public string KSCHL { get; set; }
        public string KBETR { get; set; }
        public string KONWA { get; set; }
        public int DATAB { get; set; }
        public int DATBI { get; set; }
    }
    public class ALaCartePriceListResponse
    {
        public List<LTTABLEResponse> LT_TABLE { get; set; }
    }
    public class ALaCartePriceListRequest
    {
        public List<LTTABLE> data { get; set; }
        public string IM_PROJECT { get; set; }
    }

}