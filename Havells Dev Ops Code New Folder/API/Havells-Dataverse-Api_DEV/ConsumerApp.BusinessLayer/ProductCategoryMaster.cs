// Pass Define Master = 3 for Product Category
// Pass Define Master = 4 for Product Sub-Category

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
    public class ProductCategorySubCategoryMaster
    {
        [DataMember(IsRequired = false)]
        public string PRODCAT_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string PRODCAT_ID { get; set; }
        public List<ProductCategorySubCategoryMaster> GetProductCategory()
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            List <ProductCategorySubCategoryMaster> obj = new List<ProductCategorySubCategoryMaster>();
            ColumnSet Col = new ColumnSet(true);
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = Product.EntityLogicalName;
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("hil_hierarchylevel", ConditionOperator.Equal, 2));
            EntityCollection Found = service.RetrieveMultiple(Qry);
            foreach(Product et in Found.Entities)
            {
                if(et.Name != null)
                {
                    obj.Add(
                    new ProductCategorySubCategoryMaster
                    {
                        PRODCAT_NAME = et.Name,
                        PRODCAT_ID = Convert.ToString(et.Id)//Staging MG Division Mapping Id
                    });
                }
            }
            return (obj);
        }
    }
}
