using Havells.Dataverse.CustomConnector.Consumer_Profile;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Customer_Asset
{
    public class IoTProducts : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            List<ProductDTO> lstProduct = new List<ProductDTO>();

            string productSubCategoryId = Convert.ToString(context.InputParameters["productSubCategoryId"]);
            if (string.IsNullOrWhiteSpace(productSubCategoryId))
            {
                lstProduct.Add(new ProductDTO { StatusCode = "204", StatusDescription = "productSubCategoryId is required" });
                string jsonResult = JsonSerializer.Serialize(lstProduct);
                context.OutputParameters["data"] = jsonResult;
                return;
            }
            productSubCategoryId = context.InputParameters["productSubCategoryId"].ToString();
            Guid productSubCategoryIdGuid;

            if (Guid.TryParse(productSubCategoryId, out productSubCategoryIdGuid))
            {
                if (APValidate.IsvalidGuid(productSubCategoryId))
                {
                    List<ProductDTO> result = GetProducts(productSubCategoryIdGuid, service);
                    string jsonResult = JsonSerializer.Serialize(result);
                    context.OutputParameters["data"] = jsonResult;
                }
            }
            else
            {
                lstProduct.Add(new ProductDTO { StatusCode = "204", StatusDescription = " productSubCategory GUID is incorrect" });
                string jsonResult = JsonSerializer.Serialize(lstProduct);
                context.OutputParameters["data"] = jsonResult;
                return;
            }

        }
        public List<ProductDTO> GetProducts(Guid productSubCategoryId, IOrganizationService service)
        {
            List<ProductDTO> lstProduct = new List<ProductDTO>();
            try
            {
                string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
                "<entity name='product'>" +
                "<attribute name='name' />" +
                "<attribute name='description' />" +
                "<order attribute='name' descending='false' />" +
                "<filter type='and'>" +
                    "<condition attribute='hil_materialgroup' operator='eq' value='{" + productSubCategoryId + @"}' />" +
                "</filter>" +
                "</entity>" +
                "</fetch>";

                EntityCollection entcoll;
                entcoll = service.RetrieveMultiple(new FetchExpression(fetchXML));

                if (entcoll.Entities.Count > 0)
                {
                    foreach (Entity ent in entcoll.Entities)
                    {
                        lstProduct.Add(new ProductDTO() { ProductCode = ent.GetAttributeValue<string>("name"), Product = ent.GetAttributeValue<string>("description"), ProductGuid = ent.Id, StatusCode = "200", StatusDescription = "OK" });
                    }
                }
                else
                {
                    lstProduct.Add(new ProductDTO { StatusCode = "204", StatusDescription = "Product is not found against this ProductSubCategoryId" });
                }
            }
            catch (Exception ex)
            {
                lstProduct.Add(new ProductDTO { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() });
            }
            return lstProduct;
        }
    }
}
