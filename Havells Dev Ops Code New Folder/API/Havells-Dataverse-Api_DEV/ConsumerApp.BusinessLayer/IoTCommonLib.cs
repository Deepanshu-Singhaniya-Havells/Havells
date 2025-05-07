using System.Collections.Generic;
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
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class IoTCommonLib
    {
        [DataMember]
        public Guid ProductSubCatgguId { get; set; }
        public List<HashTableDTO> GetMIMETypes()
        {
            List<HashTableDTO> retObj = new List<HashTableDTO>();
            retObj.Add(new HashTableDTO() { Label = "image/jpeg", Value = 0, Extension = ".jpeg" });
            retObj.Add(new HashTableDTO() { Label = "image/png", Value = 1, Extension = ".png" });
            retObj.Add(new HashTableDTO() { Label = "application/pdf", Value = 2, Extension = ".pdf" });
            retObj.Add(new HashTableDTO() { Label = "application/msword", Value = 3, Extension = ".doc" });
            retObj.Add(new HashTableDTO() { Label = "image/tiff", Value = 4, Extension = ".tiff" });
            retObj.Add(new HashTableDTO() { Label = "image/gif", Value = 5, Extension = ".gif" });
            retObj.Add(new HashTableDTO() { Label = "image/bmp", Value = 6, Extension = ".bmp" });
            retObj.Add(new HashTableDTO() { Label = "application/vnd.ms-excel", Value = 7, Extension = ".xls" });
            return retObj;
        }

        public List<HashTableDTO> GetSalutationEnum()
        {
            HashTableDTO objIoTSalutationEnum;
            List<HashTableDTO> lstIoTSalutationEnum = new List<HashTableDTO>();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    var attributeRequest = new RetrieveAttributeRequest
                    {
                        EntityLogicalName = "contact",
                        LogicalName = "hil_salutation",
                        RetrieveAsIfPublished = true
                    };

                    var attributeResponse = (RetrieveAttributeResponse)service.Execute(attributeRequest);
                    var attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

                    var optionList = (from o in attributeMetadata.OptionSet.Options
                                      select new { Value = o.Value, Text = o.Label.UserLocalizedLabel.Label }).ToList();
                    foreach (var option in optionList)
                    {
                        lstIoTSalutationEnum.Add(new HashTableDTO() { Value = option.Value, Label = option.Text, StatusCode = "200", StatusDescription = "OK" });
                    }
                }
                else
                {
                    objIoTSalutationEnum = new HashTableDTO { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    lstIoTSalutationEnum.Add(objIoTSalutationEnum);
                }
            }
            catch (Exception ex)
            {
                objIoTSalutationEnum = new HashTableDTO { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                lstIoTSalutationEnum.Add(objIoTSalutationEnum);
            }
            return lstIoTSalutationEnum;
        }

        public List<HashTableDTO> GetInstallationLocationEnum()
        {
            HashTableDTO objIoTSalutationEnum;
            List<HashTableDTO> lstIoTSalutationEnum = new List<HashTableDTO>();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    var attributeRequest = new RetrieveAttributeRequest
                    {
                        EntityLogicalName = "msdyn_customerasset",
                        LogicalName = "hil_product",
                        RetrieveAsIfPublished = true
                    };

                    var attributeResponse = (RetrieveAttributeResponse)service.Execute(attributeRequest);
                    var attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

                    var optionList = (from o in attributeMetadata.OptionSet.Options
                                      select new { Value = o.Value, Text = o.Label.UserLocalizedLabel.Label }).ToList();
                    foreach (var option in optionList)
                    {
                        lstIoTSalutationEnum.Add(new HashTableDTO() { Value = option.Value, Label = option.Text, StatusCode = "200", StatusDescription = "OK" });
                    }
                }
                else
                {
                    objIoTSalutationEnum = new HashTableDTO { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    lstIoTSalutationEnum.Add(objIoTSalutationEnum);
                }
            }
            catch (Exception ex)
            {
                objIoTSalutationEnum = new HashTableDTO { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                lstIoTSalutationEnum.Add(objIoTSalutationEnum);
            }
            return lstIoTSalutationEnum;
        }

        public List<ProductHierarchyDTO> GetProductHierarchy()
        {
            List<ProductHierarchyDTO> lstProductHierarchy = new List<ProductHierarchyDTO>();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
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
                    "<filter type='and'>" +
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

                        //***changed by Saurabh
                        //if (ent.Attributes.Contains("prodCatg.name"))
                        //{
                        //    objProductHierarchy.ProductCategory = ent.GetAttributeValue<AliasedValue>("prodCatg.name").Value.ToString();
                        //}
                        //objProductHierarchy.ProductCategoryGuid = ent.GetAttributeValue<EntityReference>("hil_productcategorydivision").Id;

                        objProductHierarchy.ProductCategory = ent.GetAttributeValue<EntityReference>("hil_productcategorydivision").Name;
                        objProductHierarchy.ProductCategoryGuid = ent.GetAttributeValue<EntityReference>("hil_productcategorydivision").Id;


                        //***changed by Saurabh ends here.

                        if (ent.Attributes.Contains("prodSubcatg.name"))
                        {
                            objProductHierarchy.ProductSubCategory = ent.GetAttributeValue<AliasedValue>("prodSubcatg.name").Value.ToString();
                        }
                        objProductHierarchy.ProductSubCategoryGuid = ent.GetAttributeValue<EntityReference>("hil_productsubcategorymg").Id;

                        //***changed by Saurabh
                        //if (ent.Attributes.Contains("prodSubcatg.hil_isverificationrequired"))
                        //{
                        //    objProductHierarchy.IsVerificationrequired = Convert.ToBoolean(ent.GetAttributeValue<AliasedValue>("prodSubcatg.hil_isverificationrequired").Value);
                        //}
                        //else
                        //{
                        //    objProductHierarchy.IsVerificationrequired = false;
                        //}

                        objProductHierarchy.IsVerificationrequired = ent.Attributes.Contains("prodSubcatg.hil_isverificationrequired") ? Convert.ToBoolean(ent.GetAttributeValue<AliasedValue>("prodSubcatg.hil_isverificationrequired").Value) : false;


                        //if (ent.Attributes.Contains("prodSubcatg.hil_isserialized"))
                        //{
                        //    int value = ((OptionSetValue)(ent.GetAttributeValue<AliasedValue>("prodSubcatg.hil_isserialized").Value)).Value;
                        //    objProductHierarchy.IsSerialized = value == 1 ? true : false;
                        //}
                        //else
                        //{
                        //    objProductHierarchy.IsSerialized = false;
                        //}

                        objProductHierarchy.IsSerialized = ent.Attributes.Contains("prodSubcatg.hil_isserialized") ? ((OptionSetValue)(ent.GetAttributeValue<AliasedValue>("prodSubcatg.hil_isserialized").Value)).Value == 1 ? true : false : false;

                        //***changed by Saurabh ends here.

                        objProductHierarchy.StatusCode = "200";
                        objProductHierarchy.StatusDescription = "OK";
                        lstProductHierarchy.Add(objProductHierarchy);
                    }
                }
                else
                {
                    lstProductHierarchy.Add(new ProductHierarchyDTO { StatusCode = "503", StatusDescription = "D365 Service Unavailable" });
                }
            }
            catch (Exception ex)
            {
                lstProductHierarchy.Add(new ProductHierarchyDTO { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() });
            }
            return lstProductHierarchy;
        }

        public List<ProductDTO> GetProducts(IoTCommonLib ProductSubCatgInfo)
        {
            List<ProductDTO> lstProduct = new List<ProductDTO>();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (ProductSubCatgInfo.ProductSubCatgguId == Guid.Empty)
                    {
                        lstProduct.Add(new ProductDTO { StatusCode = "204", StatusDescription = "Product Subcategory is required." });
                    }
                    else
                    {
                        string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
                        "<entity name='product'>" +
                        "<attribute name='name' />" +
                        "<attribute name='description' />" +
                        "<order attribute='name' descending='false' />" +
                        "<filter type='and'>" +
                            "<condition attribute='hil_materialgroup' operator='eq' value='{" + ProductSubCatgInfo.ProductSubCatgguId + @"}' />" +
                        "</filter>" +
                        "</entity>" +
                        "</fetch>";

                        EntityCollection entcoll;
                        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXML));

                        foreach (Entity ent in entcoll.Entities)
                        {
                            lstProduct.Add(new ProductDTO() { ProductCode = ent.GetAttributeValue<string>("name"), Product = ent.GetAttributeValue<string>("description"), ProductGuid = ent.Id, StatusCode = "200", StatusDescription = "OK" });
                        }
                    }
                }
                else
                {
                    lstProduct.Add(new ProductDTO { StatusCode = "503", StatusDescription = "D365 Service Unavailable" });
                }
            }
            catch (Exception ex)
            {
                lstProduct.Add(new ProductDTO { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() });
            }
            return lstProduct;
        }

        public Attachment AttachNotes(Attachment attachmentData)
        {
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                string title = string.Empty;
                string fileName = string.Empty;
                string mimeType = string.Empty;
                string extension = string.Empty;

                if (attachmentData.Base64String == null)
                {
                    return new Attachment() { StatusCode = "204", StatusDescription = "No Content! Base64String is required." };
                }
                if (attachmentData.ObjectType == null)
                {
                    return new Attachment() { StatusCode = "204", StatusDescription = "No Content! Object Type is required." };
                }
                if (attachmentData.ObjectGuid == Guid.Empty)
                {
                    return new Attachment() { StatusCode = "204", StatusDescription = "No Content! Object GuId is required." };
                }
                if (attachmentData.FileType == null)
                {
                    return new Attachment() { StatusCode = "204", StatusDescription = "No Content! File Type is required." };
                }
                if (attachmentData.FileType < 0 || attachmentData.FileType > 7)
                {
                    return new Attachment() { StatusCode = "204", StatusDescription = "No Content! File Type is not within File Type Range." };
                }
                if (attachmentData.ImageType == null)
                {
                    return new Attachment() { StatusCode = "204", StatusDescription = "No Content! Image Type is required." };
                }

                List<HashTableDTO> FileTypeList = new List<HashTableDTO>();
                FileTypeList = this.GetMIMETypes();
                mimeType = FileTypeList.FirstOrDefault(x => x.Value == attachmentData.FileType).Label;
                extension = FileTypeList.FirstOrDefault(x => x.Value == attachmentData.FileType).Extension;

                if (attachmentData.ImageType == null)
                {
                    return new Attachment() { StatusCode = "204", StatusDescription = "No Content! Invalid File Type." };
                }

                if (attachmentData.ImageType == ImageType.InvoiceImage)
                {
                    fileName = attachmentData.ObjectGuid.ToString() + "InvoiceImage" + extension;
                }

                else if (attachmentData.ImageType == ImageType.InvoiceImage)
                {
                    fileName = attachmentData.ObjectGuid.ToString() + "ProductImage" + extension;
                }

                byte[] NoteByte = Convert.FromBase64String(attachmentData.Base64String);
                Annotation An = new Annotation();
                An.DocumentBody = Convert.ToBase64String(NoteByte);
                An.MimeType = mimeType;
                An.FileName = fileName;
                An.ObjectId = new EntityReference(attachmentData.ObjectType, attachmentData.ObjectGuid);
                An.ObjectTypeCode = attachmentData.ObjectType;
                service.Create(An);
                return new Attachment() { StatusCode = "200", StatusDescription = "OK" };
            }
            catch (Exception ex)
            {
                return new Attachment() { StatusCode = "204", StatusDescription = "Error!!!." + ex.Message };
            }
        }
    }

    [DataContract]
    public enum ImageType
    {
        InvoiceImage = 0,
        ProductImage = 1
    }

    [DataContract]
    public class Attachment
    {
        [DataMember]
        public string ObjectType { get; set; }
        [DataMember]
        public Guid ObjectGuid { get; set; }
        [DataMember]
        public string Base64String { get; set; }
        [DataMember]
        public ImageType? ImageType { get; set; }
        [DataMember]
        public int? FileType { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }
}
