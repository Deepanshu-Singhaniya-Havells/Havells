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
    public class ProductCodeMaster
    {
        [DataMember]
        public string PRODSUBCAT_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string PROD_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string PROD_ID { get; set; }
        public List<ProductCodeMaster> GetProductMaster(ProductCodeMaster iMaster)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            List<ProductCodeMaster> obj = new List<ProductCodeMaster>();
            ColumnSet Col = new ColumnSet(true);
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = Product.EntityLogicalName;
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            Qry.Criteria.AddCondition(new ConditionExpression("hil_hierarchylevel", ConditionOperator.Equal, 5));
            Qry.Criteria.AddCondition(new ConditionExpression("hil_materialgroup", ConditionOperator.Equal, iMaster.PRODSUBCAT_ID));
            EntityCollection Found = service.RetrieveMultiple(Qry);
            foreach (Product et in Found.Entities)
            {
                if (et.Name != null)
                {
                    obj.Add(
                    new ProductCodeMaster
                    {
                        PROD_NAME = et.Name,
                        PROD_ID = Convert.ToString(et.Id)//Staging MG Division Mapping Id
                    });
                }
            }
            return (obj);
        }
    }
}
