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
    public class ProductSubCategoryMaster
    {
        [DataMember]
        public string PRODCAT_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string PRODSUBCAT_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string PRODSUBCAT_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string PRODSUBCATMAPPING_ID { get; set; }
        public List<ProductSubCategoryMaster> GetProductSubCategory(ProductSubCategoryMaster iMaster)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            List<ProductSubCategoryMaster> obj = new List<ProductSubCategoryMaster>();
            try
            {
                ColumnSet Col = new ColumnSet(true);
                QueryExpression Qry = new QueryExpression();
                Qry.EntityName = hil_stagingdivisonmaterialgroupmapping.EntityLogicalName;
                Qry.ColumnSet = Col;
                Qry.Criteria = new FilterExpression(LogicalOperator.And);
                Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                Qry.Criteria.AddCondition(new ConditionExpression("hil_productcategorydivision", ConditionOperator.Equal, new Guid(iMaster.PRODCAT_ID)));
                Qry.Criteria.AddCondition(new ConditionExpression("hil_productsubcategorymg", ConditionOperator.NotNull));
                Qry.Criteria.AddCondition(new ConditionExpression("hil_productsubcategorymg", ConditionOperator.NotEqual, new Guid("4AABBB57-A85E-EA11-A811-000D3AF057DD")));
                EntityCollection Found = service.RetrieveMultiple(Qry);
                foreach (hil_stagingdivisonmaterialgroupmapping et in Found.Entities)
                {
                    if (et.hil_name != null)
                    {
                        obj.Add(
                        new ProductSubCategoryMaster
                        {
                            PRODSUBCAT_NAME = et.hil_name,
                            PRODSUBCATMAPPING_ID = Convert.ToString(et.Id),//Staging MG Division Mapping Id
                            PRODSUBCAT_ID = et.hil_ProductSubCategoryMG.Id.ToString()
                        });
                    }
                }
            }
            catch(Exception ex)
            {
                obj.Add
                    (
                        new ProductSubCategoryMaster
                        {
                            PRODSUBCAT_NAME = ex.Message.ToUpper(),
                            PRODSUBCATMAPPING_ID = ex.Message.ToUpper(),//Staging MG Division Mapping Id
                            PRODSUBCAT_ID = ex.Message.ToUpper()
                        });
            }
            return (obj);
        }
    }
}
