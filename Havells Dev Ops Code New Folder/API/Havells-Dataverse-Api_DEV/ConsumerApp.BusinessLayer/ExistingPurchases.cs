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
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class ExisitingPurchases
    {
        [DataMember]
        public string CustomerGuId { get; set; }
        [DataMember(IsRequired = false)]
        public string SerialNumber { get; set; }
        [DataMember(IsRequired = false)]
        public string ProductItem { get; set; }
        [DataMember(IsRequired = false)]
        public string ProductCategory { get; set; }
        [DataMember(IsRequired = false)]
        public string InvoiceNumber { get; set; }
        [DataMember(IsRequired = false)]
        public DateTime InvoiceDate { get; set; }
        [DataMember(IsRequired = false)]
        public DateTime WarrantyDate { get; set; }
        [DataMember(IsRequired = false)]
        public string ProductSubCategory { get; set; }
        [DataMember(IsRequired = false)]
        public string AssetId { get; set; }
        [DataMember(IsRequired = false)]
        public string CAT_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string SCAT_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string PROD_INSTALL { get; set; }
        [DataMember(IsRequired = false)]
        public string PRODUCT_DESCRIPTION { get; set; }
        [DataMember(IsRequired = false)]
        public string CustAssetStatus { get; set; }
        [DataMember(IsRequired = false)]
        public string Address { get; set; }
        [DataMember(IsRequired = false)]
        public string Scheme { get; set; }
    
        public List<ExisitingPurchases> GetAllExistingPurchase(ExisitingPurchases Purchase)
        {
            List<ExisitingPurchases> obj = new List<ExisitingPurchases>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = "msdyn_customerasset";
            ColumnSet Col = new ColumnSet(true);
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("hil_customer", ConditionOperator.Equal, new Guid(Purchase.CustomerGuId)));
            Qry.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.NotNull));
            Qry.AddOrder("createdon", OrderType.Descending);
            EntityCollection Colec = service.RetrieveMultiple(Qry);
            foreach (msdyn_customerasset et in Colec.Entities)
            {
                EntityReference PdtCtgry = new EntityReference("product");
                EntityReference PdtItem = new EntityReference("product");
                EntityReference PSubCat = new EntityReference("product");
                Entity Pdt = new Entity("product");
                Entity ProdtCtgry = new Entity("product");
                string Item = string.Empty;
                string ItemCat = string.Empty;
                string InvNo = string.Empty;
                string SerialNumber = string.Empty;
                string ProdSubCat = string.Empty;
                string Asset_Id = string.Empty;
                string Item_Id = string.Empty;
                string ItemSubCatId = string.Empty;
                string ProductInstallLoc = string.Empty;
                string ProductDescription = string.Empty;
                string CustomerAssetstatus = string.Empty;
                string CustAddress = string.Empty;
                Asset_Id = et.Id.ToString();
                DateTime InvDate = new DateTime();
                DateTime WarDate = new DateTime();
                if (et.Attributes.Contains("hil_productcategory"))
                {
                    PdtCtgry = (EntityReference)et["hil_productcategory"];
                    //ProdtCtgry = service.Retrieve("product", PdtCtgry.Id, new ColumnSet("name"));
                    ItemCat = PdtCtgry.Name;
                    Item_Id = PdtCtgry.Id.ToString();
                }
                if (et.Attributes.Contains("msdyn_product"))
                {
                    PdtItem = (EntityReference)et["msdyn_product"];
                    //Pdt = service.Retrieve("product", PdtItem.Id, new ColumnSet("name"));
                    Item = PdtItem.Name;
                }
                if (et.Attributes.Contains("hil_invoiceno"))
                {
                    InvNo = Convert.ToString(et["hil_invoiceno"]);
                }
                if (et.Attributes.Contains("hil_invoicedate"))
                {
                    InvDate = Convert.ToDateTime(et["hil_invoicedate"]);
                }
                if (et.Attributes.Contains("hil_warrantytilldate"))
                {
                    WarDate = Convert.ToDateTime(et["hil_warrantytilldate"]);
                }
                if (et.Attributes.Contains("msdyn_name"))
                {
                    SerialNumber = Convert.ToString(et["msdyn_name"]);
                }
              
                if (et.Attributes.Contains("hil_productsubcategory"))
                {
                    PSubCat = (EntityReference)et["hil_productsubcategory"];
                    ProdSubCat = PSubCat.Name;
                    ItemSubCatId = PSubCat.Id.ToString();
                }
                if (et.Attributes.Contains("hil_product"))
                {
                    ProductInstallLoc = et.FormattedValues["hil_product"];
                    //OptionSetValue opProdLoc = (OptionSetValue)et["hil_product"];
                    ////Living Room -910590009
                    ////Balcony - 910590000
                    ////Kitchen - 910590008
                    ////Bed Room 1 - 910590004
                    ////Bed Room 2 - 910590005
                    ////Bed Room 3 - 910590006
                    ////Bed Room 4 - 910590007
                    ////Bath Room 1 - 910590001
                    ////Bath Room 2 - 910590002
                    ////Bath Room 3 - 910590003
                    ////Lobby - 910590010
                    ////Office - 910590011
                    ////Parking - 910590012
                    ////Terrace - 910590013
                    ////Others - 910590014
                    //if (opProdLoc.Value == 910590009)
                    //{
                    //    ProductInstallLoc = "Living Room";
                    //}
                    //else if (opProdLoc.Value == 910590000)
                    //    ProductInstallLoc = "Balcony";
                    //else if (opProdLoc.Value == 910590004)
                    //    ProductInstallLoc = "Bed Room 1";
                    //else if (opProdLoc.Value == 910590005)
                    //    ProductInstallLoc = "Bed Room 2";
                    //else if (opProdLoc.Value == 910590006)
                    //    ProductInstallLoc = "Bed Room 3";
                    //else if (opProdLoc.Value == 910590008)
                    //    ProductInstallLoc = "Kitchen";
                    //else if (opProdLoc.Value == 910590007)
                    //    ProductInstallLoc = "Bed Room 4";
                    //else if (opProdLoc.Value == 910590001)
                    //    ProductInstallLoc = "Bath Room 1";
                    //else if (opProdLoc.Value == 910590002)
                    //    ProductInstallLoc = "Bath Room 2";
                    //else if (opProdLoc.Value == 910590003)
                    //    ProductInstallLoc = "Bath Room 3";
                    //else if (opProdLoc.Value == 910590010)
                    //    ProductInstallLoc = "Lobby";
                    //else if (opProdLoc.Value == 910590011)
                    //    ProductInstallLoc = "Office";
                    //else if (opProdLoc.Value == 910590012)
                    //    ProductInstallLoc = "Parking";
                    //else if (opProdLoc.Value == 910590013)
                    //    ProductInstallLoc = "Terrace";
                    //else if (opProdLoc.Value == 910590014)
                    //    ProductInstallLoc = "Others";
                    //else
                    //    ProductInstallLoc = "";
                }
                if (et.Contains("hil_modelname"))
                {
                    ProductDescription = et.hil_ModelName;
                }
                if (et.Contains("hil_branchheadapprovalstatus"))
                {
                    CustomerAssetstatus = et.FormattedValues["hil_branchheadapprovalstatus"];
                }
                if (et.Contains("hil_pincode"))
                {
                    QueryExpression Query = new QueryExpression(hil_address.EntityLogicalName);
                    Query.ColumnSet = new ColumnSet("hil_pincode");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, Purchase.CustomerGuId);
                    Query.Criteria.AddCondition("hil_businessgeo", ConditionOperator.Equal, et.GetAttributeValue<EntityReference>("hil_pincode").Id);
                    EntityCollection Found = service.RetrieveMultiple(Query);
                    if (Found.Entities.Count > 0)
                    {
                        hil_address iAddress = Found.Entities[0].ToEntity<hil_address>();
                        CustAddress = iAddress.hil_FullAddress != null ? iAddress.hil_FullAddress : string.Empty;
                    }
                }
                obj.Add(
                new ExisitingPurchases
                {
                    SerialNumber = SerialNumber,
                    InvoiceNumber = InvNo,
                    InvoiceDate = InvDate,
                    WarrantyDate = WarDate,
                    ProductCategory = ItemCat,
                    ProductSubCategory = ProdSubCat,
                    ProductItem = Item,
                    AssetId = Asset_Id,
                    CAT_ID = Item_Id,
                    SCAT_ID = ItemSubCatId,
                    PROD_INSTALL = ProductInstallLoc,
                    PRODUCT_DESCRIPTION = ProductDescription,
                    CustAssetStatus = CustomerAssetstatus,
                    Address = CustAddress
                });
            }
            return (obj);
        }
    }
}
