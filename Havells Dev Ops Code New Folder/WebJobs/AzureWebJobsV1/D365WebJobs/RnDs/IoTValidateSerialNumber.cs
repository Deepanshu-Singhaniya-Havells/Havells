using System;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Runtime.Serialization.Json;
using System.IO;
using Havells_Plugin.HelperIntegration;

namespace D365WebJobs
{
    
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

        public IoTValidateSerialNumber ValidateAssetSerialNumber(IoTValidateSerialNumber _reqParam, IOrganizationService service)
        {
            IoTValidateSerialNumber retObj;
            try
            {
                //IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (_reqParam.SerialNumber == string.Empty || _reqParam.SerialNumber.Trim().Length == 0)
                    {
                        return new IoTValidateSerialNumber { StatusCode = "204", StatusDescription = "Asset Serial Number is required." };
                    }
                    retObj = new IoTValidateSerialNumber() { StatusCode = "204", StatusDescription = "Something went wrong." };
                    //bool IfExisting = CheckIfExistingSerialNumber(service, _reqParam.SerialNumber);
                    Entity entAsset = CheckIfExistingSerialNumberWithDetails(service, _reqParam.SerialNumber);

                    if (entAsset == null)
                    {
                        using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                        {
                            #region Credentials
                            String sUserName = String.Empty;
                            String sPassword = String.Empty;
                            var obj2 = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                                       where _IConfig.hil_name == "Credentials"
                                       select new { _IConfig };
                            foreach (var iobj2 in obj2)
                            {
                                if (iobj2._IConfig.hil_Username != String.Empty)
                                    sUserName = iobj2._IConfig.hil_Username;
                                if (iobj2._IConfig.hil_Password != String.Empty)
                                    sPassword = iobj2._IConfig.hil_Password;
                            }
                            #endregion
                            var obj = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                                      where _IConfig.hil_name == "SerialNumberValidation"
                                      select new { _IConfig.hil_Url };
                            foreach (var iobj in obj)
                            {

                                if (iobj.hil_Url != null)
                                {
                                    String sUrl = iobj.hil_Url + _reqParam.SerialNumber;
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
                                        var obj1 = (from _Product in orgContext.CreateQuery<Product>()
                                                    where _Product.ProductNumber == rootObject.EX_PRD_DET.MATNR
                                                    select new
                                                    {
                                                        _Product.ProductId,
                                                        _Product.hil_MaterialGroup,
                                                        _Product.hil_Division,
                                                        _Product.ProductNumber,
                                                        _Product.Description
                                                    }).Take(1);
                                        foreach (var iobj1 in obj1)
                                        {
                                            if (iobj1.hil_Division == null | iobj1.hil_MaterialGroup == null) {
                                                return new IoTValidateSerialNumber()
                                                {
                                                    SerialNumber = _reqParam.SerialNumber,
                                                    StatusCode = "204",
                                                    StatusDescription = "Division or Material Group mapping is missing."
                                                };
                                            }
                                            retObj = new IoTValidateSerialNumber()
                                            {
                                                ProductId = iobj1.ProductId,
                                                ModelCode = iobj1.ProductNumber,
                                                ModelName = iobj1.Description,
                                                ProductCategory = iobj1.hil_Division.Name,
                                                ProductCategoryGuid = iobj1.hil_Division.Id,
                                                ProductSubCategory = iobj1.hil_MaterialGroup.Name,
                                                ProductSubCategoryGuid = iobj1.hil_MaterialGroup.Id,
                                                SerialNumber = _reqParam.SerialNumber,
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
                                            SerialNumber = _reqParam.SerialNumber,
                                            StatusCode = "204",
                                            StatusDescription = "SERIAL NUMBER NOT VALID"
                                        };
                                    }
                                }
                                else
                                {
                                    retObj = new IoTValidateSerialNumber()
                                    {
                                        SerialNumber = _reqParam.SerialNumber,
                                        StatusCode = "204",
                                        StatusDescription = "SAP API Config not found."
                                    };
                                }
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

        public bool CheckIfExistingSerialNumber(IOrganizationService service, string Serial)
        {
            bool IfExisting = true;
            QueryExpression Query = new QueryExpression(msdyn_customerasset.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, Serial);
            EntityCollection Found = service.RetrieveMultiple(Query);
            {
                if (Found.Entities.Count > 0)
                {
                    IfExisting = true;
                }
                else
                {
                    IfExisting = false;
                }
            }
            return IfExisting;
        }

        public Entity CheckIfExistingSerialNumberWithDetails(IOrganizationService service, string Serial)
        {
            Entity entAsset = null;
            QueryExpression Query = new QueryExpression(msdyn_customerasset.EntityLogicalName);
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

        public IoTValidateSerialNumber GetProductDetails(IoTValidateSerialNumber _reqParam, IOrganizationService service)
        {
            IoTValidateSerialNumber retObj = null;
            try
            {
                //IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (_reqParam.ModelCode == string.Empty || _reqParam.ModelCode.Trim().Length == 0)
                    {
                        retObj = new IoTValidateSerialNumber { StatusCode = "204", StatusDescription = "Model Number is required." };
                    }
                    else
                    {
                        using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                        {
                            #region MaterialCodeOfAsset

                            QueryExpression Query = new QueryExpression("product");
                            Query.ColumnSet = new ColumnSet("name", "productnumber", "description", "hil_materialgroup", "hil_division");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("name", ConditionOperator.Equal, _reqParam.ModelCode);
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
                            }
                            else
                            {
                                retObj = new IoTValidateSerialNumber()
                                {
                                    StatusCode = "204",
                                    StatusDescription = "Model Number does not exist."
                                };
                            }
                            #endregion
                        }
                    }
                }
                else
                {
                    retObj = new IoTValidateSerialNumber { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                retObj = new IoTValidateSerialNumber { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return retObj;
        }
    }
}
