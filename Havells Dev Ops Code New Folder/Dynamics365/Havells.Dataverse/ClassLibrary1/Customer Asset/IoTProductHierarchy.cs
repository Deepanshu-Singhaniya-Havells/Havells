using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Customer_Asset
{
    public class IoTProductHierarchy : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            List<ProductHierarchyDTO> result = GetProductHierarchy(service);

            string jsonResult = JsonSerializer.Serialize(result);
            context.OutputParameters["data"] = jsonResult;
        }
        public List<ProductHierarchyDTO> GetProductHierarchy(IOrganizationService service)
        {
            List<ProductHierarchyDTO> lstProductHierarchy = new List<ProductHierarchyDTO>();
            try
            {
                string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
                "<entity name='hil_stagingdivisonmaterialgroupmapping'>" +
                "<attribute name='hil_productcategorydivision' />" +
                "<attribute name='hil_productsubcategorymg' />" +
                "<order attribute='hil_productcategorydivision' descending='false' />" +
                "<link-entity name='product' from='productid' to='hil_productsubcategorymg' visible='false' link-type='outer' alias='prodSubcatg'>" +
                "<attribute name='hil_isverificationrequired' />" +
                "<attribute name='hil_isserialized' />" +
                "<attribute name='name' />" +
                "</link-entity>" +
                "<link-entity name='product' from='productid' to='hil_productcategorydivision' visible='false' link-type='outer' alias='prodCatg'>" +
                "<attribute name='name' />" +
                "</link-entity>" +
                "<filter type='and'>" +
                    "<condition attribute='statecode' operator='eq' value='0' />" +
                    "<condition attribute='hil_productcategorydivision' operator='not-null' />" +
                    "<condition attribute='hil_productsubcategorymg' operator='not-null' />" +
                "</filter>" +
                "</entity>" +
                "</fetch>";
                EntityCollection entcoll;
                entcoll = service.RetrieveMultiple(new FetchExpression(fetchXML));

                foreach (Entity ent in entcoll.Entities)
                {
                    ProductHierarchyDTO objProductHierarchy = new ProductHierarchyDTO();
                    if (ent.Attributes.Contains("prodCatg.name"))
                    {
                        objProductHierarchy.ProductCategory = ent.GetAttributeValue<AliasedValue>("prodCatg.name").Value.ToString();
                    }
                    objProductHierarchy.ProductCategoryGuid = ent.GetAttributeValue<EntityReference>("hil_productcategorydivision").Id;
                    if (ent.Attributes.Contains("prodSubcatg.name"))
                    {
                        objProductHierarchy.ProductSubCategory = ent.GetAttributeValue<AliasedValue>("prodSubcatg.name").Value.ToString();
                    }
                    objProductHierarchy.ProductSubCategoryGuid = ent.GetAttributeValue<EntityReference>("hil_productsubcategorymg").Id;
                    if (ent.Attributes.Contains("prodSubcatg.hil_isverificationrequired"))
                    {
                        objProductHierarchy.IsVerificationrequired = Convert.ToBoolean(ent.GetAttributeValue<AliasedValue>("prodSubcatg.hil_isverificationrequired").Value);
                    }
                    else
                    {
                        objProductHierarchy.IsVerificationrequired = false;
                    }
                    if (ent.Attributes.Contains("prodSubcatg.hil_isserialized"))
                    {
                        int value = ((OptionSetValue)(ent.GetAttributeValue<AliasedValue>("prodSubcatg.hil_isserialized").Value)).Value;
                        objProductHierarchy.IsSerialized = value == 1 ? true : false;
                    }
                    else
                    {
                        objProductHierarchy.IsSerialized = false;
                    }
                    objProductHierarchy.StatusCode = "200";
                    objProductHierarchy.StatusDescription = "OK";
                    lstProductHierarchy.Add(objProductHierarchy);
                }
            }
            catch (Exception ex)
            {
                lstProductHierarchy.Add(new ProductHierarchyDTO { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() });
            }
            return lstProductHierarchy;
        }
        public class ProductHierarchyDTO
        {
            public string ProductCategory { get; set; }
            public Guid ProductCategoryGuid { get; set; }
            public string ProductSubCategory { get; set; }
            public Guid ProductSubCategoryGuid { get; set; }
            public bool IsSerialized { get; set; }
            public bool IsVerificationrequired { get; set; }
            public string StatusCode { get; set; }
            public string StatusDescription { get; set; }
        }
    }
}
