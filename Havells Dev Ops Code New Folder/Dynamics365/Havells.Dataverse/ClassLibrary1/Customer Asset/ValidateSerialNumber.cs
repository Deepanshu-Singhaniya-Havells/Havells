using Havells.Dataverse.CustomConnector.Consumer_GeoLocations;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace Havells.Dataverse.CustomConnector.Customer_Asset
{
    public class ValidateSerialNumber : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            string JsonResponse = "";
            string SerialNumber = null;
            try
            {


                if (context.InputParameters.Contains("SerialNumber") && context.InputParameters["SerialNumber"] is string)
                {
                    SerialNumber = Convert.ToString(context.InputParameters["SerialNumber"]).Trim();

                    if (string.IsNullOrWhiteSpace(SerialNumber) || !APValidate.isAlphaNumeric(SerialNumber))
                    {
                        JsonResponse = JsonSerializer.Serialize(new IoTValidateSerialNumber
                        {
                            StatusCode = "204",
                            StatusDescription = "Invalid Serial Number."
                        });
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                    JsonResponse = JsonSerializer.Serialize(ValidateAssetSerialNumber(service, SerialNumber));
                    _tracingService.Trace(JsonResponse);
                    context.OutputParameters["data"] = JsonResponse;
                }
                else
                {
                    JsonResponse = JsonSerializer.Serialize(new IoTValidateSerialNumber
                    {
                        StatusCode = "204",
                        StatusDescription = "Serial Number is required."
                    });
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
            }
            catch (Exception ex)
            {
                var retObj = JsonSerializer.Serialize(new IoTValidateSerialNumber { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() });
                context.OutputParameters["data"] = retObj;
                return;
            }
        }
        public IoTValidateSerialNumber ValidateAssetSerialNumber(IOrganizationService service, string SerialNumber)
        {
            IoTValidateSerialNumber retObj;
            try
            {
                if (service != null)
                {
                    retObj = new IoTValidateSerialNumber() { StatusCode = "204", StatusDescription = "Something went wrong." };
                    Entity entAsset = CheckIfExistingSerialNumberWithDetails(service, SerialNumber);

                    if (entAsset == null)
                    {
                        using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                        {
                            #region Credentials
                            QueryExpression Query = new QueryExpression("hil_integrationconfiguration");
                            Query.ColumnSet = new ColumnSet("hil_username", "hil_password");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "Credentials");
                            EntityCollection EntColl = service.RetrieveMultiple(Query);
                            string sUserName = EntColl.Entities[0].Contains("hil_username") ? EntColl.Entities[0].GetAttributeValue<string>("hil_username") : null;
                            string sPassword = EntColl.Entities[0].Contains("hil_password") ? EntColl.Entities[0].GetAttributeValue<string>("hil_password") : null;
                            #endregion
                            Query = new QueryExpression("hil_integrationconfiguration");
                            Query.ColumnSet = new ColumnSet("hil_url");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "SerialNumberValidation");
                            EntColl = service.RetrieveMultiple(Query);

                            string hil_Url = EntColl.Entities[0].Contains("hil_url") ? EntColl.Entities[0].GetAttributeValue<string>("hil_url") : null;

                            if (hil_Url != null)
                            {
                                String sUrl = hil_Url + SerialNumber;
                                WebClient webClient = new WebClient();
                                string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
                                webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;

                                webClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                                var jsonData = webClient.DownloadData(sUrl);
                                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(SerialNumberValidation));
                                SerialNumberValidation rootObject = (SerialNumberValidation)ser.ReadObject(new MemoryStream(jsonData));
                                if (rootObject.EX_PRD_DET != null)//valid
                                {
                                    #region MaterialCodeOfAsset
                                    string fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' top='1' distinct='false'>
                                                        <entity name='product'>
                                                        <attribute name='name'/>
                                                        <attribute name='productnumber'/>
                                                        <attribute name='description'/>
                                                        <attribute name='productid'/>
                                                        <attribute name='hil_productcode'/>
                                                        <attribute name='hil_materialgroup'/>
                                                        <attribute name='hil_division'/>
                                                        <order attribute='productnumber' descending='false'/>
                                                        <filter type='and'>
                                                            <condition attribute='name' operator='eq' value='{rootObject.EX_PRD_DET.MATNR}' />
                                                        </filter>
                                                        </entity>
                                                        </fetch>";
                                    EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(fetchQuery));
                                    if (_entCol.Entities.Count > 0)
                                    {
                                        retObj = new IoTValidateSerialNumber()
                                        {
                                            ProductId = _entCol.Entities[0].Id,
                                            ModelCode = _entCol.Entities[0].Contains("productnumber") ? _entCol.Entities[0].GetAttributeValue<string>("productnumber") : "",
                                            ModelName = _entCol.Entities[0].Contains("description") ? _entCol.Entities[0].GetAttributeValue<string>("description") : "",
                                            ProductCategory = _entCol.Entities[0].Contains("hil_division") ? _entCol.Entities[0].GetAttributeValue<EntityReference>("hil_division").Name : "",
                                            ProductCategoryGuid = _entCol.Entities[0].Contains("hil_division") ? _entCol.Entities[0].GetAttributeValue<EntityReference>("hil_division").Id : Guid.Empty,
                                            ProductSubCategory = _entCol.Entities[0].Contains("hil_materialgroup") ? _entCol.Entities[0].GetAttributeValue<EntityReference>("hil_materialgroup").Name : "",
                                            ProductSubCategoryGuid = _entCol.Entities[0].Contains("hil_materialgroup") ? _entCol.Entities[0].GetAttributeValue<EntityReference>("hil_materialgroup").Id : Guid.Empty,
                                            SerialNumber = SerialNumber,
                                            StatusCode = "200",
                                            StatusDescription = "OK"
                                        };
                                    }
                                    #endregion
                                }
                                else//Not valid
                                {
                                    retObj = new IoTValidateSerialNumber()
                                    {
                                        SerialNumber = SerialNumber,
                                        StatusCode = "204",
                                        StatusDescription = "Invalid Serial Number."
                                    };
                                }
                            }
                            else
                            {
                                retObj = new IoTValidateSerialNumber()
                                {
                                    SerialNumber = SerialNumber,
                                    StatusCode = "204",
                                    StatusDescription = "SAP API Config not found."
                                };
                            }
                        }
                        return retObj;
                    }
                    else
                    {
                        retObj = new IoTValidateSerialNumber() { StatusCode = "204", StatusDescription = "Serial Number already exist." };
                        return retObj;
                    }
                }
                else
                {
                    retObj = new IoTValidateSerialNumber { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    return retObj;
                }
            }
            catch (Exception ex)
            {
                retObj = new IoTValidateSerialNumber { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                return retObj;
            }
        }
        public Entity CheckIfExistingSerialNumberWithDetails(IOrganizationService service, string Serial)
        {
            Entity entAsset = null;
            QueryExpression Query = new QueryExpression("msdyn_customerasset");
            Query.ColumnSet = new ColumnSet("msdyn_name");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, Serial);
            EntityCollection Found = service.RetrieveMultiple(Query);
            {
                if (Found.Entities.Count > 0)
                {
                    entAsset = Found.Entities[0];
                }
            }
            return entAsset;
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
    public class SerialNumberValidation
    {
        public EX_PRD_DET EX_PRD_DET;
    }
    public class EX_PRD_DET
    {
        public string SERIAL_NO;
        public string MATNR;
        public string MAKTX;
        public string SPART;
        public string REGIO;
        public string VBELN;
        public string FKDAT;
        public string KUNAG;
        public string NAME1;
        public string MFD;
        public string WTY_STATUS;
    }
}
