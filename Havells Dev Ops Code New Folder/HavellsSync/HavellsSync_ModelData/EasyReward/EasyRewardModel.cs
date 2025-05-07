using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Web.UI.WebControls;

namespace HavellsSync_ModelData.EasyReward
{

    public class UserinfoDetails
    {
        public string FirstName { get; set; }

        public string LastName { get; set; }

        public string Gender { get; set; }
        public string EmailId { get; set; }
        public string DOB { get; set; }
        public char LoyaltyMember { get; set; } = 'N';
        public char Eligible { get; set; } = 'N';

        public Guid Guid { get; set; }
        public string PreferredLanguage { get; set; }

        public int StatusCode { get; set; }


    }

    public class Product
    {
        public string ModelCode { get; set; }

        public string Category { get; set; }

        public decimal InvoiceAmount { get; set; }

        public string InvoiceDate { get; set; }

        public string SerialNumber { get; set; }

    }

    public class EasyRewardResponse
    {

        public string Response { get; set; }

    }
    public class Loyaltyparam<T>
    {
        public T data { get; set; }
    }
    public class LoyaltySourceType
    {
       
        public string SourceType { get; set; }
       
    }

    public class SKUBillDetails
    {
        public string CountryCode { get; set; }
        public string TransactionDate { get; set; }
        public string BillNo { get; set; }
        public string Channel { get; set; }
        public string CustomerType { get; set; }
        public decimal BillValue { get; set; }
        public string PointsRedeemed { get; set; }
        public string PointsValueRedeemed { get; set; }
        public string PrimaryOrderNumber { get; set; }
        public string ShippingCharges { get; set; }
        public string PreDelivery { get; set; }
        public string SKUOfferCode { get; set; }
        public TransactionItems TransactionItems { get; set; }
        public PaymentMode PaymentMode { get; set; }

    }

    public class TransactionItem
    {
        public string ItemType { get; set; }
        public string ItemQty { get; set; }
        public string ItemName { get; set; }
        public string Size { get; set; }
        public string Unit { get; set; }
        public string ItemDiscount { get; set; }
        public string ItemTax { get; set; }
        public decimal TotalPrice { get; set; }
        public decimal BilledPrice { get; set; }
        public string Department { get; set; }
        public string Category { get; set; }
        public string Group { get; set; }
        public string ItemId { get; set; }
        public string RefBillNo { get; set; }
    }

    public class TransactionItems
    {
        public List<TransactionItem> TransactionItem { get; set; }
    }

    public class TenderItem
    {
        public string TenderCode { get; set; }
        public string TenderID { get; set; }
        public decimal TenderValue { get; set; }
    }

    public class PaymentMode
    {
        public List<TenderItem> TenderItem { get; set; }
    }

}






