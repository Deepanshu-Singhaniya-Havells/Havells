using HavellsSync_ModelData.AMC;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace HavellsSync_ModelData.Consumer
{



    public class Product
    {
        public string ModelCode { get; set; }

        public string Category { get; set; }

        public decimal InvoiceAmount { get; set; }

        public string InvoiceDate { get; set; }

        public string SerialNumber { get; set; }

    }

    public class ConsumerResponse
    {

        public string Response { get; set; }

    }
    public class Consumerparam<T>
    {
        public T data { get; set; }
    }
    public class ConsumerSourceType
    {
        public string SourceType { get; set; }

        public string Rating { get; set; }
        public string Review { get; set; }

    }

    ///////////////////// Invoice Details /////////////////


    public class InvoiceDetailsResponse
    {
        public string OrderNumber { get; set; }
        public string SAPInvoiceNumber { get; set; }
        public string SAPOrderNumber { get; set; }
        public string SAPSyncMessage { get; set; }
    }

    public class InvoiceResponse
    {
        public string Response { get; set; }
        public object Data { set; get; }

    }
    public class Invoiceparamdata<T>
    {
        public T Data { get; set; }
    }
    public class Invoiceparam
    {
        public string FromDate { get; set; }

        public string ToDate { get; set; }
        public string OrderNumber { get; set; }

    }
    public class PriceListParam
    {
        public string MATNR { get; set; }
        public string KSCHL { get; set; }
        public string KBETR { get; set; }
        public string KONWA { get; set; }
        public string DATAB { get; set; }
        public string DATBI { get; set; }

    }

    public class OCLDetailsParam
    {
        public string OrderNumber { get; set; }
    }

    public class OCLDetailsResponse
    {
        public string OCLNumber { get; set; }
        public string ApprovedDataSheet { get; set; }
        public string DrumLengthSchedule { get; set; }
        public string StdDrumLength { get; set; }
        public string QtyTolerance { get; set; }
        public string Inspection { get; set; }
        public string MarkingDetails { get; set; }
        public string TypeOfDrum { get; set; } 


    }

    #region MFRServiceJobs
    public class WorkOrderResponse
    {

        public List<WorkOrderInfo> workOrderInfos { get; set; }

    }

    public class WorkOrderInfo
    {
        public string JobId { get; set; }

        public string Substatus { get; set; }

        public string Createdon { get; set; }

        public string Productsubcategory { get; set; }

        public string Productcategory { get; set; }

        public string Callsubtype { get; set; }

        public string Customer { get; set; }

        public string Owner { get; set; }

    }

    public class WorkOrderRequest
    {
        public string DealerCode { get; set; }
        public string FromDate { get; set; }
        public string ToDate { get; set; }
    }

    public class JobStatusDTO
    {

        public string job_id { get; set; }
        public string mobile_number { get; set; }
        public string serial_number { get; set; }
        public string product_category { get; set; }
        public string product_subcategory { get; set; }
        public string call_type { get; set; }
        public string customer_complaint { get; set; }
        public string assigned_resource { get; set; }
        public string job_substatus { get; set; }
        public string technician_remarks { get; set; }
        public string closed_on { get; set; }
        public string cancel_reason { get; set; }
        public string closure_remarks { get; set; }
        public string webclosure_remarks { get; set; }
        public List<JobProductDTO> spare_parts { get; set; }
        public string status_code { get; set; }
        public string status_description { get; set; }
    }

    public class JobProductDTO
    {
        public string index { get; set; }
        public string product_code { get; set; }
        public string product_description { get; set; }
    }

    public class JobRequestDTO
    {
        public string customer_firstname { get; set; }

        public string customer_lastname { get; set; }

        public string customer_mobileno { get; set; }

        public string alternate_number { get; set; }

        public string address_line1 { get; set; }

        public string address_line2 { get; set; }

        public string landmark { get; set; }

        public string pincode { get; set; }

        public string call_type { get; set; }

        public string product_subcategory { get; set; }

        public string caller_type { get; set; }

        public string dealer_code { get; set; }

        public string expected_delivery_date { get; set; }

        public string chief_complaint { get; set; }

        public string job_id { get; set; }

        public string status_code { get; set; }

        public string status_description { get; set; }

    }

    #endregion


    #region WCF Services

    public class JobOutput
    {
        public string Job_ID { get; set; }
        public string MobileNumber { get; set; }
        public string CallSubType { get; set; }
        public string Job_Loggedon { get; set; }
        public string Job_Status { get; set; }
        public string Job_AssignedTo { get; set; }
        public string Job_Asset { get; set; }
        public string Job_Category { get; set; }
        public string Job_NOC { get; set; }
        public string Job_ClosedOn { get; set; }
        public string Customer_name { get; set; }
        public string Customer_Address { get; set; }
        public string Product { get; set; }
        public Guid ProductCategoryGuid { get; set; }
        public string ProductCategoryName { get; set; }
        public string ChiefComplaint { get; set; }
    }

    public class Job
    {

        public string Job_ID { get; set; }
        public string MobileNumber { get; set; }
    }

    public class IoTServiceCallResult
    {
         public string JobId { get; set; }
        public Guid JobGuid { get; set; }
        public string CallSubType { get; set; }
        public string JobLoggedon { get; set; }
        public string JobStatus { get; set; }
        public string JobAssignedTo { get; set; }
        public string CustomerAsset { get; set; }
        public string ProductCategory { get; set; }
        public string NatureOfComplaint { get; set; }
        public string JobClosedOn { get; set; }
        public string CustomerName { get; set; }
        public string ServiceAddress { get; set; }
        public string Product { get; set; }
        public string ChiefComplaint { get; set; }
        public string PreferredDate { get; set; }
        public int PreferredPartOfDay { get; set; }
        public string PreferredPartOfDayName { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
    }

    public class IotServiceCall
    {
        public Guid CustomerGuid { get; set; }
    }

    public class IoTRegisteredProducts
    {
        public Guid CustomerGuid { get; set; }
        public Guid RegisteredProductGuid { get; set; }
        public string ProductCategory { get; set; }
        public Guid ProductCategoryId { get; set; }
        public string ProductSubCategory { get; set; }
        public Guid ProductSubCategoryId { get; set; }
        public string DealerPinCode { get; set; }
        public Guid ProductId { get; set; }
        public string ProductCode { get; set; }
        public string ProductName { get; set; }
        public string SerialNumber { get; set; }
        public string BatchNumber { get; set; }
        public bool InvoiceAvailable { get; set; }
        public string InvoiceNumber { get; set; }
        public string InvoiceDate { get; set; }
        public decimal? InvoiceValue { get; set; }
        public string PurchasedFrom { get; set; }
        public string PurchasedFromLocation { get; set; }
        public string InstalledLocation { get; set; }
        public int InstalledLocationEnum { get; set; }
        public int SourceOfRegistration { get; set; }
        public string WarrantyStatus { get; set; }
        public string WarrantySubStatus { get; set; }
        public string WarrantyEndDate { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
        public List<IoTProductWarranty> ProductWarranty { get; set; }
    }

    public class IoTProductWarranty
    {
        public string WarrantyType { get; set; }
        public string WarrantyStartDate { get; set; }
        public string WarrantyEndDate { get; set; }
        public string WarrantySpecifications { get; set; }
    }

    public class IoT_RegisterConsumer
    {
        public string MobileNumber { get; set; }
        public string Email { get; set; }
        public string FirstName { get; set; }
        public int? Salutation { get; set; }
        public int? SourceOfCreation { get; set; }
        public bool? Consent { get; set; }
        public bool? SubscribeForMsgService { get; set; }
        public string PreferredLanguage { get; set; }


    }
    public class ReturnResult
    {
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
        public Guid CustomerGuid { get; set; }
        public string MobileNumber { get; set; }
        public string CustomerName { get; set; }
        public string EmailId { get; set; }
        public bool? Consent { get; set; }
        public bool? SubscribeForMsgService { get; set; }
        public string PreferredLanguage { get; set; }
        public string PINCode { get; set; }
        public string Address { get; set; }
    }

    public class HashTableDTO
    {
        public string Label { get; set; }       
        public int? Value { get; set; }       
        public string Extension { get; set; }       
        public string StatusCode { get; set; }       
        public string StatusDescription { get; set; }
    }

    public class IoTNatureofComplaint
    {       
        public string SerialNumber { get; set; }       
        public Guid ProductSubCategoryId { get; set; }       
        public string Name { get; set; }       
        public Guid Guid { get; set; }       
        public string StatusCode { get; set; }       
        public string StatusDescription { get; set; }      
        public string Source { get; set; }
    }

    public class NatureOfComplaint
    {       
        public string SerialNumber { get; set; }       
        public string Name { get; set; }       
        public string ProductCategoryName { get; set; }       
        public Guid Guid { get; set; }       
        public Guid ProductCategoryGuid { get; set; }       
        public bool ResultStatus { get; set; }       
        public string ResultMessage { get; set; }
    }

        #endregion
    }







