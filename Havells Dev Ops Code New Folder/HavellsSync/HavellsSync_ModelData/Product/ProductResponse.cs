using HavellsSync_ModelData.Common;
using System;
using System.Collections.Generic;

namespace HavellsSync_ModelData.Product
{

    public class ProductResponse
    {
        public string ProductName { get; set; }
        public string ProductDescription { get; set; } = "";
        public string Status { get; set; }
    }
    public class ProductNotesResponse
    {
        public bool Status { get; set; }
        public string Message { get; set; }
    }
    public class SerialNumberModel
    {
        public string SerialNumber { get; set; }
    }

    public class ValidateSerialNumber : TokenExpires
    {
        public string SerialNumber { get; set; }
        public string ProductCategory { get; set; }
        public Guid ProductCategoryGuid { get; set; }
        public string ProductSubCategory { get; set; }
        public Guid ProductSubCategoryGuid { get; set; }
        public string ModelCode { get; set; }
        public string ModelName { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
        public Guid? ProductId { get; set; }
    }
    public class ValidateSerialNumResponse: TokenExpires
    {
        public string ModelCode { get; set; }
        public string ModelName { get; set; }
        public string ProductCategory { get; set; }
        public Guid ProductCategoryGuid { get; set; }
        public string ProductId { get; set; }
        public string ProductSubCategory { get; set; }
        public Guid ProductSubCategoryGuid { get; set; }
        public string SerialNumber { get; set; }
        public string Brand { get; set; }
        public bool PTAStatus { get; set; }
        public int StatusCode { get; set; }
        public string Message { get; set; } 
    }
    public class RegisterProductModel : TokenExpires
    {
        public Guid ProductId { get; set; }
        public string SerialNumber { get; set; } = null;
        public string InvoiceDate { get; set; } = null;
        public string InvoiceBase64 { get; set; } = null;
        public int FileType { get; set; }
        public Guid CustomerGuid { get; set; }
        public int SourceType { get; set; } 

    }
    public class RegisterProductResponse : TokenExpires
    {
        public string Brand { get; set; }
        public string ProductDesc { get; set; }
        public Guid CustomerAssetGuid { get; set; }
        public int StatusCode { get; set; }
        public string Message { get; set; }

    }
    public class ProductWarranty
    {
        public string WarrantyEndDate { get; set; }
        public string WarrantySpecifications { get; set; }
        public string WarrantyStartDate { get; set; }
        public string WarrantyType { get; set; }
    }
    public class ProductDeatilsList : TokenExpires
    {
        public List<ProductDetails> ProductList { get; set; }
        public int StatusCode { get; set; }
        public string Message { get; set; }

    }
    public class ProductDetails
    {
        public Guid CustomerGuid { get; set; }
        public string Brand { get; set; }
        public bool InvoiceAvailable { get; set; }
        public string InvoiceDate { get; set; }
        public string InvoiceNumber { get; set; }
        public decimal InvoiceValue { get; set; }
        public string ProductCategory { get; set; }
        public Guid ProductCategoryId { get; set; }
        public Guid ProductId { get; set; }
        public string ModelCode { get; set; }
        public string ProductDesc { get; set; }
        public string ProductSubCategory { get; set; }
        public Guid ProductSubCategoryId { get; set; }
        public List<ProductWarranty> ProductWarranty { get; set; }
        public string PurchasedFrom { get; set; }
        public string PurchasedFromLocation { get; set; }
        public Guid ProductGuid { get; set; }
        public string SerialNumber { get; set; }
        public string SourceOfRegistration { get; set; }
        public string WarrantyEndDate { get; set; }
        public string WarrantyStatus { get; set; }
        public string WarrantySubStatus { get; set; }

    }
    public class SerialNumberValidation
    {
        public EX_PRD_DET EX_PRD_DET { get; set; }
    }
    public class EX_PRD_DET
    {

        public string SERIAL_NO;


        public string MATNR;


        public string MAKTX;


        public string SPART;


        public string REGIO;


        public string VBELN;


        public string FKDAT;


        public string KUNAG;


        public string NAME1;


        public string WTY_STATUS;
    }
}