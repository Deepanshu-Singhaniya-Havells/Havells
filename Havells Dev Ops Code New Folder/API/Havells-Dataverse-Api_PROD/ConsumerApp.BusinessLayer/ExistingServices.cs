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
using Microsoft.Xrm.Sdk.Deployment;
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class ExistingServices
    {
        [DataMember]
        public string CustomerGuId { get; set; }
        [DataMember(IsRequired = false)]
        public string WoID { get; set; }
        [DataMember(IsRequired = false)]
        public string ProductItem { get; set; }
        [DataMember(IsRequired = false)]
        public string ProductCategory { get; set; }
        [DataMember(IsRequired = false)]
        public string RaisedOn { get; set; }
        [DataMember(IsRequired = false)]
        public string ProductSubCategory { get; set; }
        [DataMember(IsRequired = false)]
        public string CallSubtype { get; set; }
        [DataMember(IsRequired = false)]
        public string Description { get; set; }
        [DataMember(IsRequired = false)]
        public string CurrentStatus { get; set; }

        [DataMember(IsRequired = false)]
        public string CloseOn { get; set; }

        [DataMember(IsRequired = false)]
        public string SerialNo { get; set; }

        [DataMember(IsRequired = false)]
        public string InvoiceNo { get; set; }

        [DataMember(IsRequired = false)]
        public string InvoiceDate { get; set; }
        [DataMember(IsRequired = false)]
        public string WarrantyDate { get; set; }

        [DataMember(IsRequired = false)]
        public string Location { get; set; }
        [DataMember(IsRequired = false)]
        public string Address { get; set; }
        [DataMember(IsRequired = false)]
        public string ComplaintDescription { get; set; }

        public List<ExistingServices> GetAllExistingServices(ExistingServices Service)
        {
            List<ExistingServices> obj = new List<ExistingServices>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = "msdyn_workorder";
            ColumnSet Col = new ColumnSet(true);
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("hil_customerref", ConditionOperator.Equal,
                new Guid(Service.CustomerGuId)));
            Qry.AddOrder("msdyn_name", OrderType.Descending);
            EntityCollection Colec = service.RetrieveMultiple(Qry);
            foreach (msdyn_workorder et in Colec.Entities)
            {
                EntityReference PdtCtgry = new EntityReference("product");
                EntityReference PdtItem = new EntityReference("product");
                EntityReference PdtSubCat = new EntityReference("product");
                string PdtSubCatName = string.Empty;
                Entity Pdt = new Entity("product");
                Entity ProdtCtgry = new Entity("product");
                string Item = string.Empty;
                string ItemCat = string.Empty;
                string InvNo = string.Empty;
                string CallSTypName = string.Empty;
                string Desc = string.Empty;
                string iSubStatus = string.Empty;
                DateTime RaisedOn = new DateTime();
                DateTime ClosedOn = new DateTime();
                string i_serialno = string.Empty;
                string i_invoiceno = string.Empty;
                string i_invoicedate = string.Empty;
                string i_warrantydate = string.Empty;
                string i_location = string.Empty;
                string i_description = string.Empty;
                if (et.Attributes.Contains("hil_productcategory"))
                {
                    PdtCtgry = (EntityReference)et["hil_productcategory"];
                    //ProdtCtgry = service.Retrieve("product", PdtCtgry.Id, new ColumnSet("name"));
                    ItemCat = PdtCtgry.Name;
                }
                //if (et.Attributes.Contains("hil_productsubcategory"))
                //{
                //    PdtItem = (EntityReference)et["hil_productsubcategory"];
                //    Item = PdtItem.Name;
                //}
                if (et.Attributes.Contains("hil_jobclosureon"))
                {
                    ClosedOn = (DateTime)et["hil_jobclosureon"];
                }
                if (et.Attributes.Contains("createdon"))
                {
                    RaisedOn = (DateTime)et["createdon"];
                }
                if (et.Attributes.Contains("hil_productsubcategory"))
                {
                    PdtSubCat = (EntityReference)et["hil_productsubcategory"];
                    PdtSubCatName = PdtSubCat.Name;
                }
                if (et.Attributes.Contains("hil_callsubtype"))
                {
                    EntityReference Callstyp = (EntityReference)et["hil_callsubtype"];
                    CallSTypName = Callstyp.Name;
                }
                if (et.Attributes.Contains("msdyn_substatus"))
                {
                    EntityReference SubStatus = (EntityReference)et["msdyn_substatus"];
                    iSubStatus = SubStatus.Name;
                }

                if (et.msdyn_CustomerAsset != null)
                {
                    msdyn_customerasset Cust_Asset = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName,
                        et.msdyn_CustomerAsset.Id, new ColumnSet(true));

                    i_serialno = Cust_Asset.msdyn_name;
                    i_description = Cust_Asset.hil_ModelName;
                    i_invoiceno = Cust_Asset.hil_InvoiceNo != null ? Cust_Asset.hil_InvoiceNo : string.Empty;
                    Item = Cust_Asset.msdyn_Product != null ? Cust_Asset.msdyn_Product.Name : string.Empty;
                    i_invoicedate = Cust_Asset.hil_InvoiceDate != null ? Cust_Asset.hil_InvoiceDate.ToString() : string.Empty;
                    if (Cust_Asset.Attributes.ContainsKey("hil_warrantytilldate"))
                        i_warrantydate = Cust_Asset.Attributes["hil_warrantytilldate"].ToString();
                    i_location = Cust_Asset.hil_Product != null ? Cust_Asset.FormattedValues["hil_product"] : string.Empty;
                }
                obj.Add(
                new ExistingServices
                {
                    WoID = Convert.ToString(et["msdyn_name"]),
                    ComplaintDescription = Convert.ToString(et.hil_CustomerComplaintDescription != null ?
                   et.hil_CustomerComplaintDescription : string.Empty),
                    RaisedOn = Convert.ToString(RaisedOn),
                    Description = i_description,
                    ProductCategory = ItemCat,
                    ProductSubCategory = PdtSubCatName,
                    ProductItem = Item,
                    CallSubtype = CallSTypName,
                    CurrentStatus = iSubStatus,
                    CloseOn = Convert.ToString(ClosedOn),
                    SerialNo = i_serialno,
                    InvoiceNo = i_invoiceno,
                    InvoiceDate = i_invoicedate,
                    WarrantyDate = i_warrantydate,
                    Location = i_location,
                    Address = et.hil_FullAddress != null ? et.hil_FullAddress : string.Empty
                });
            }
            return (obj);
        }
    }
}
