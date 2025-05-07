using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Runtime.Serialization;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class FindMaterialByName
    {
        [DataMember]
        public string MATNR { get; set; }
        [DataMember(IsRequired = false)]
        public string ModelGuid { get; set; }
        [DataMember(IsRequired = false)]
        public string ModelName { get; set; }
        [DataMember(IsRequired = false)]
        public string SubCategoryId { get; set; }
        [DataMember(IsRequired = false)]
        public string SubCategoryName { get; set; }
        [DataMember(IsRequired = false)]
        public string CategoryId { get; set; }
        [DataMember(IsRequired = false)]
        public string CategoryName { get; set; }
        public FindMaterialByName GetMaterialId(FindMaterialByName Pf)
        {
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                QueryExpression Query = new QueryExpression(Product.EntityLogicalName);
                Query.ColumnSet = new ColumnSet("name", "description", "hil_materialgroup", "hil_division");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("name", ConditionOperator.Equal, Pf.MATNR);
                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    Pf.ModelGuid = Found.Entities[0].Id.ToString();
                    Pf.ModelName = Found.Entities[0].ToEntity<Product>().Description;
                    Pf.SubCategoryId = Found.Entities[0].ToEntity<Product>().hil_MaterialGroup.Id.ToString();
                    Pf.SubCategoryName = Found.Entities[0].ToEntity<Product>().hil_MaterialGroup.Name.ToString();
                    Pf.CategoryId = Found.Entities[0].ToEntity<Product>().hil_Division.Id.ToString();
                    Pf.CategoryName = Found.Entities[0].ToEntity<Product>().hil_Division.Name.ToString();
                }
            }
            catch(Exception ex)
            {
                Pf.ModelGuid = ex.Message.ToUpper();
                Pf.ModelName = ex.Message.ToUpper();
                Pf.SubCategoryId = " ";
                Pf.SubCategoryName = "";
                Pf.CategoryId = " ";
                Pf.CategoryName = " ";
            }
            return(Pf);
        }
    }
}