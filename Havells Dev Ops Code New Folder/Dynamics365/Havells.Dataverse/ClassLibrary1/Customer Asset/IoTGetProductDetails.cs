using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;

namespace Havells.Dataverse.CustomConnector.Customer_Asset
{
    public class IoTGetProductDetails : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            string JsonResponse = "";
            IoTValidateSerialNumber retObj = null;
            try
            {
                
                string ModelCode = Convert.ToString(context.InputParameters["ModelCode"]);

                if (service != null)
                {
                    if (string.IsNullOrWhiteSpace(ModelCode))
                    {
                        retObj = new IoTValidateSerialNumber { StatusCode = "204", StatusDescription = "Model Code is required." };
                        JsonResponse = JsonConvert.SerializeObject(retObj, Formatting.Indented);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                    else
                    {
                        using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                        {
                            #region MaterialCodeOfAsset

                            QueryExpression Query = new QueryExpression("product");
                            Query.ColumnSet = new ColumnSet("name", "productnumber", "description", "hil_materialgroup", "hil_division");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("name", ConditionOperator.Equal, ModelCode);
                            EntityCollection Found = service.RetrieveMultiple(Query);
                            if (Found.Entities.Count > 0)
                            {
                                retObj = new IoTValidateSerialNumber();
                                retObj.ProductId = Found.Entities[0].Id;
                                if (Found.Entities[0].Attributes.Contains("name"))
                                {
                                    retObj.ModelCode = Found.Entities[0].GetAttributeValue<string>("name");
                                }
                                if (Found.Entities[0].Attributes.Contains("description"))
                                {
                                    retObj.ModelName = Found.Entities[0].GetAttributeValue<string>("description");
                                }
                                if (Found.Entities[0].Attributes.Contains("hil_materialgroup"))
                                {
                                    retObj.ProductSubCategoryGuid = Found.Entities[0].GetAttributeValue<EntityReference>("hil_materialgroup").Id;
                                    retObj.ProductSubCategory = Found.Entities[0].GetAttributeValue<EntityReference>("hil_materialgroup").Name;
                                }
                                if (Found.Entities[0].Attributes.Contains("hil_division"))
                                {
                                    retObj.ProductCategoryGuid = Found.Entities[0].GetAttributeValue<EntityReference>("hil_division").Id;
                                    retObj.ProductCategory = Found.Entities[0].GetAttributeValue<EntityReference>("hil_division").Name;
                                }
                                retObj.StatusCode = "200";
                                retObj.StatusDescription = "Ok";

                                JsonResponse = JsonConvert.SerializeObject(retObj, Formatting.Indented);
                                context.OutputParameters["data"] = JsonResponse;
                                return;
                            }
                            else
                            {
                                retObj = new IoTValidateSerialNumber()
                                {
                                    StatusCode = "204",
                                    StatusDescription = "Model Code does not exist."
                                };
                                JsonResponse = JsonConvert.SerializeObject(retObj, Formatting.Indented);
                                context.OutputParameters["data"] = JsonResponse;
                                return;
                            }
                            #endregion
                        }
                    }
                }
                else
                {
                    retObj = new IoTValidateSerialNumber { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    JsonResponse = JsonConvert.SerializeObject(retObj, Formatting.Indented);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
            }
            catch (Exception ex)
            {
                retObj = new IoTValidateSerialNumber { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                JsonResponse = JsonConvert.SerializeObject(retObj, Formatting.Indented);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }

        }

        public class IoTValidateSerialNumber
        {
            public string SerialNumber { get; set; }
            public string ProductCategory { get; set; }
            public Guid ProductCategoryGuid { get; set; }
            public string ProductSubCategory { get; set; }
            public Guid ProductSubCategoryGuid { get; set; }
            public string ModelCode { get; set; }
            public string ModelName { get; set; }
            public string StatusCode { get; set; }
            public string StatusDescription { get; set; }
            public Guid? ProductId { get; set; }
        }

    }
}