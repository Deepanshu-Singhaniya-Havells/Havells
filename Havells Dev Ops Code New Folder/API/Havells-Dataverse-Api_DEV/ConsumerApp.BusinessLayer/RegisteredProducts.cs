using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class RegisteredProducts
    {
        [DataMember]
        public string SerialNumber { get; set; }
        [DataMember]
        public string ProductCategory { get; set; }
        [DataMember]
        public string ProductSubCategory { get; set; }
        [DataMember]
        public string DealerPinCode { get; set; }
        [DataMember]
        public string ProductCode { get; set; }
        [DataMember]
        public string ProductName { get; set; }
        [DataMember]
        public Guid CustomerGuid { get; set; }
        [DataMember]
        public string CustomerMobleNo { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }

        public List<RegisteredProducts> GetRegisteredProducts(RegisteredProducts registeredProduct)
        {
            Guid customerGuid = Guid.Empty;
            RegisteredProducts objRegisteredProducts;
            List<RegisteredProducts> lstRegisteredProducts = new List<RegisteredProducts>();
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (registeredProduct.CustomerMobleNo.Trim().Length == 0)
                {
                    objRegisteredProducts = new RegisteredProducts { ResultStatus = false, ResultMessage = "Mobile No. is required." };
                    lstRegisteredProducts.Add(objRegisteredProducts);
                    return lstRegisteredProducts;
                }

                Query = new QueryExpression("contact");
                Query.ColumnSet = new ColumnSet("fullname");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, registeredProduct.CustomerMobleNo);
                entcoll = service.RetrieveMultiple(Query);

                if (entcoll.Entities.Count == 0)
                {
                    objRegisteredProducts = new RegisteredProducts { ResultStatus = false, ResultMessage = "Mobile No. does not exist." };
                    lstRegisteredProducts.Add(objRegisteredProducts);
                    return lstRegisteredProducts;
                }
                else
                {
                    customerGuid = entcoll.Entities[0].Id;
                    string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_customerasset'>
                            <attribute name='createdon' />
                            <attribute name='msdyn_product' />
                            <attribute name='msdyn_name' />
                            <attribute name='hil_productsubcategorymapping' />
                            <attribute name='hil_productcategory' />
                            <attribute name='msdyn_customerassetid' />
                            <attribute name='hil_pincode' />
                            <attribute name='hil_modelname' />
                            <order attribute='msdyn_name' descending='false' />
	                        <filter type='and'>
                                <condition attribute='hil_customer' operator='eq' value='{" + customerGuid.ToString() + @"}' />
                            </filter>
                            <link-entity name='hil_businessmapping' from='hil_businessmappingid' to='hil_pincode' visible='false' link-type='outer' alias='pincode'>
                                <attribute name='hil_pincode' />
                            </link-entity>
                        </entity>
                        </fetch>";
                    entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (entcoll.Entities.Count == 0)
                    {
                        objRegisteredProducts = new RegisteredProducts { ResultStatus = false, ResultMessage = "No Product registered with Mobile No."};
                        lstRegisteredProducts.Add(objRegisteredProducts);
                    }
                    else {
                        foreach (Entity ent in entcoll.Entities) {
                            objRegisteredProducts = new RegisteredProducts();
                            objRegisteredProducts.CustomerGuid = ent.GetAttributeValue<Guid>("msdyn_customerassetid");
                            objRegisteredProducts.CustomerMobleNo = registeredProduct.CustomerMobleNo;
                            objRegisteredProducts.DealerPinCode = ent.GetAttributeValue<string>("pincode.hil_pincode");
                            objRegisteredProducts.ProductCategory = ent.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                            objRegisteredProducts.ProductCode = ent.GetAttributeValue<EntityReference>("msdyn_product").Name;
                            objRegisteredProducts.ProductName = ent.GetAttributeValue<string>("hil_modelname");
                            objRegisteredProducts.ProductSubCategory = ent.GetAttributeValue<EntityReference>("hil_productsubcategorymapping").Name;
                            objRegisteredProducts.SerialNumber = ent.GetAttributeValue<string>("msdyn_name");
                            objRegisteredProducts.ResultStatus = true;
                            objRegisteredProducts.ResultMessage = "Success";
                            lstRegisteredProducts.Add(objRegisteredProducts);
                        }
                    }
                    return lstRegisteredProducts;
                }
            }
            catch (Exception ex)
            {
                objRegisteredProducts = new RegisteredProducts { ResultStatus = false, ResultMessage = ex.Message };
                lstRegisteredProducts.Add(objRegisteredProducts);
                return lstRegisteredProducts;
            }
        }

    }
}
