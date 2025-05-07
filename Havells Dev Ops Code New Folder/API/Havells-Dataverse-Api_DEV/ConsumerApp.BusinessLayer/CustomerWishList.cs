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
    public class CustomerWishList
    {
        [DataMember]
        public int AppName { get; set; }
        [DataMember]
        public string CustomerGuId { get; set; }
        [DataMember]
        public int Type { get; set; }
        [DataMember(IsRequired = false)]
        public string ProductItem { get; set; }
        [DataMember(IsRequired = false)]
        public string ProductCat { get; set; }
        [DataMember(IsRequired = false)]
        public DateTime WishListAddedOn { get; set; }
        public List<CustomerWishList> GetAllCustomerWishList(CustomerWishList _WshLst)
        {
            EntityReference Pdt = new EntityReference("product");
            List<CustomerWishList> obj = new List<CustomerWishList>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = "hil_customerwishlist";
            ColumnSet Col = new ColumnSet("hil_productcode", "hil_productname", "createdon");
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("hil_customer", ConditionOperator.Equal, new Guid(_WshLst.CustomerGuId)));
            Qry.Criteria.AddCondition(new ConditionExpression("hil_appname", ConditionOperator.Equal, _WshLst.AppName));
            Qry.Criteria.AddCondition(new ConditionExpression("hil_type", ConditionOperator.Equal, _WshLst.Type));
            Qry.AddOrder("createdon", OrderType.Descending);
            EntityCollection Found = service.RetrieveMultiple(Qry);
            if(Found.Entities.Count > 0)
            {
                foreach (Entity Wsh in Found.Entities)
                {
                    if (Wsh.Attributes.Contains("hil_productcode"))
                    {
                        Pdt = (EntityReference)Wsh["hil_productcode"];
                        string PdtGrp = GetProductGroup(service, Pdt.Id);
                        obj.Add(new CustomerWishList
                        {
                            ProductItem = Convert.ToString(Wsh["hil_productname"]),
                            ProductCat = PdtGrp,
                            WishListAddedOn = Convert.ToDateTime(Wsh["createdon"])
                        });
                    }
                }
            }
            else
            {
                obj.Add(new CustomerWishList
                {
                    ProductItem = "",
                    ProductCat = "",
                    WishListAddedOn = new DateTime()
                });
            }
            
            return (obj);
        }
        public string GetProductGroup(IOrganizationService service, Guid PdtId)
        {
            string CatName = string.Empty;
            EntityReference PdtCtgry = new EntityReference("product");
            Entity Pdt = service.Retrieve("product", PdtId, new ColumnSet("hil_materialgroup"));
            if (Pdt.Attributes.Contains("hil_materialgroup"))
            {
                PdtCtgry = (EntityReference)Pdt["hil_materialgroup"];
                CatName = PdtCtgry.Name;
            }
            return (CatName);
        }
    }
}