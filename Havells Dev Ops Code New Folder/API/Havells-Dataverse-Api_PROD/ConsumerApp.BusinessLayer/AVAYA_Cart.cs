using System;
using System.Net;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using System.ServiceModel.Web;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using Microsoft.Crm.Sdk.Messages;
using System.Runtime.Serialization;

namespace ConsumerApp.BusinessLayer
{
    public class AVAYA_Cart
    {

        [DataMember(IsRequired = false)]

        public string ProductCode { get; set; }

        [DataMember(IsRequired = false)]

        public string ProductName { get; set; }

        [DataMember(IsRequired = false)]

        public string ProductDescription { get; set; }

        public ReturnInfo CreateCart(AVAYA_Cart Wsh, Guid CustId)
        {
            ReturnInfo _retIn = new ReturnInfo();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (CustId != null)
                {
                    EntityReference Product = GetThisProduct(service, Wsh.ProductCode);
                    if (Product != null)
                    {
                        hil_customerwishlist WshList = new hil_customerwishlist();
                        WshList.hil_Customer = new EntityReference(Contact.EntityLogicalName, CustId);
                        WshList.hil_Description = Wsh.ProductDescription;
                        WshList.hil_ProductName = Wsh.ProductName;
                        WshList.hil_ProductCode = Product;
                        WshList["hil_type"] = new OptionSetValue(2);
                        service.Create(WshList);
                        _retIn.CustomerGuid = CustId;
                        _retIn.ErrorCode = "SUCCESS";
                        _retIn.ErrorDescription = "";
                    }
                }
                else
                {
                    _retIn.CustomerGuid = CustId;
                    _retIn.ErrorCode = "FAILURE";
                    _retIn.ErrorDescription = "CUSTOMER CART : CUSTOMER ID NULL";
                }
            }
            catch (Exception ex)
            {
                _retIn.CustomerGuid = CustId;
                _retIn.ErrorCode = "FAILURE";
                _retIn.ErrorDescription = "CUSTOMER CART : " + ex.Message.ToUpper();
            }
            return _retIn;
        }
        public static EntityReference GetThisProduct(IOrganizationService service, string ProductCode)
        {
            EntityReference Pduct = new EntityReference();
            QueryExpression Query = new QueryExpression(Product.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("productnumber", ConditionOperator.Equal, ProductCode);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
                Pduct = new EntityReference(Product.EntityLogicalName, Found.Entities[0].Id);
            return Pduct;
        }
    }
}
