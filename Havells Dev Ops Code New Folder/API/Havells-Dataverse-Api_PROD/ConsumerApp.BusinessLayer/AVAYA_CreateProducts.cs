using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace ConsumerApp.BusinessLayer
{
    public class AVAYA_CreateProducts
    {
        [DataMember(IsRequired = false)]

        public string ProductName { get; set; }

        [DataMember(IsRequired = false)]

        public string ProductCode { get; set; }

        [DataMember(IsRequired = false)]

        public decimal Price { get; set; }

        [DataMember(IsRequired = false)]

        public int Quantity { get; set; }

        [DataMember(IsRequired = false)]

        public string InvoiceNo { get; set; }

        [DataMember(IsRequired = false)]

        public DateTime InvoiceDate { get; set; }

        [DataMember(IsRequired = false)]

        public string PaymentStatus { get; set; }

        [DataMember(IsRequired = false)]

        public string ProductStatus { get; set; }

        [DataMember(IsRequired = false)]

        public string TransactionID { get; set; }

        [DataMember(IsRequired = false)]

        public string ModeofPayment { get; set; }

        [DataMember(IsRequired = false)]

        public bool isDiscounted { get; set; }
        public ReturnInfo CreateTransaction(AVAYA_CreateProducts Pdts, Guid CustomerId)
        {
            ReturnInfo _ret = new ReturnInfo();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (Pdts.ProductCode != null)
                {
                    Guid PdtCode = GetThisProductCode(Pdts.ProductCode, service);
                    msdyn_customerasset Asset = new msdyn_customerasset();
                    if (PdtCode != Guid.Empty && CustomerId != Guid.Empty)
                    {
                        Asset.msdyn_Product = new EntityReference(Product.EntityLogicalName, PdtCode);
                        Product ThisPdt = (Product)service.Retrieve(Product.EntityLogicalName, PdtCode, new ColumnSet("hil_division", "hil_materialgroup"));
                        if (ThisPdt.hil_MaterialGroup != null)
                        {
                            Asset.hil_ProductSubcategory = ThisPdt.hil_MaterialGroup;
                        }
                        if (ThisPdt.hil_Division != null)
                        {
                            Asset.hil_ProductCategory = ThisPdt.hil_Division;
                        }
                        Asset.hil_Customer = new EntityReference(Contact.EntityLogicalName, CustomerId);
                        Guid DummyId = GetDummyAccount(service, "Dummy Customer");
                        if (DummyId != Guid.Empty)
                            Asset.msdyn_Account = new EntityReference(Account.EntityLogicalName, DummyId);
                        if (Pdts.isDiscounted != null)
                        {
                            Asset["hil_isdiscounted"] = Pdts.isDiscounted;
                        }
                        if (Pdts.ModeofPayment != null)
                        {
                            Asset["hil_modeofpayment"] = Pdts.ModeofPayment;
                        }
                        if (Pdts.TransactionID != null)
                        {
                            Asset["hil_transactionid"] = Pdts.TransactionID;
                        }
                        if (Pdts.ProductStatus != null)
                        {
                            Asset["hil_productstatus"] = Pdts.ProductStatus;
                        }
                        if (Pdts.PaymentStatus != null)
                        {
                            Asset["hil_paymentstatus"] = Pdts.PaymentStatus;
                        }
                        if (Pdts.InvoiceDate != null)
                        {
                            Asset.hil_InvoiceDate = Pdts.InvoiceDate;
                        }
                        if (Pdts.InvoiceNo != null)
                        {
                            Asset.hil_InvoiceNo = Pdts.InvoiceNo;
                        }
                        if (Pdts.Quantity > 0)
                        {
                            Asset["hil_quantity"] = Pdts.Quantity;
                        }
                        if (Pdts.Price != null)
                        {
                            Asset["hil_invoicevalue"] = Pdts.Price;
                        }
                        Guid AssetId = service.Create(Asset);
                        if(AssetId != Guid.Empty)
                        {
                            _ret.CustomerGuid = CustomerId;
                            _ret.ErrorCode = "SUCCESS";
                            _ret.ErrorDescription = "";
                        }
                    }
                    else
                    {
                        _ret.CustomerGuid = CustomerId;
                        _ret.ErrorCode = "FAILURE";
                        _ret.ErrorDescription = "UNABLE TO CREATE ASSET : PRODUCT ID EMPTY OR CUSTOMER NOT FOUND";
                    }
                }
            }
            catch(Exception ex)
            {
                _ret.CustomerGuid = CustomerId;
                _ret.ErrorCode = "FAILURE";
                _ret.ErrorDescription = "CUSTOMER ASSET : " +ex.Message.ToUpper();
            }
            return _ret;
        }
        public static Guid GetThisProductCode(string PdtCode, IOrganizationService service)
        {
            Guid ProductId = new Guid();
            QueryExpression Query = new QueryExpression(Product.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("productnumber", ConditionOperator.Equal, PdtCode);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count > 0)
            {
                ProductId = Found.Entities[0].Id;
            }
            return ProductId;
        }
        public static Guid GetDummyAccount(IOrganizationService service, string Dummy)
        {
            Guid DummyId = new Guid();
            QueryExpression Query = new QueryExpression(Account.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("name", ConditionOperator.Equal, Dummy);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                DummyId = Found.Entities[0].Id;
            }
            return DummyId;
        }
    }
}
