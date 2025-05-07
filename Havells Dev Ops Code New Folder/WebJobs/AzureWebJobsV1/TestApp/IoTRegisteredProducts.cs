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
using System.Globalization;

namespace TestApp
{
    [DataContract]
    public class IoTRegisteredProducts
    {
        [DataMember]
        public Guid CustomerGuid { get; set; }
        [DataMember]
        public Guid RegisteredProductGuid { get; set; }
        [DataMember]
        public string ProductCategory { get; set; }
        [DataMember]
        public Guid ProductCategoryId { get; set; }
        [DataMember]
        public string ProductSubCategory { get; set; }
        [DataMember]
        public Guid ProductSubCategoryId { get; set; }
        [DataMember]
        public string DealerPinCode { get; set; }
        [DataMember]
        public Guid ProductId { get; set; }
        [DataMember]
        public string ProductCode { get; set; }
        [DataMember]
        public string ProductName { get; set; }
        [DataMember]
        public string SerialNumber { get; set; }
        [DataMember]
        public string BatchNumber { get; set; }
        [DataMember]
        public bool InvoiceAvailable { get; set; }
        [DataMember]
        public string InvoiceNumber { get; set; }
        [DataMember]
        public string InvoiceDate { get; set; }
        [DataMember]
        public Decimal? InvoiceValue { get; set; }
        [DataMember]
        public string PurchasedFrom { get; set; }
        [DataMember]
        public string PurchasedFromLocation { get; set; }
        [DataMember]
        public string InstalledLocation { get; set; }
        [DataMember]
        public int InstalledLocationEnum { get; set; }
        [DataMember]
        public int SourceOfRegistration { get; set; }
        [DataMember]
        public string WarrantyStatus { get; set; }
        [DataMember]
        public string WarrantySubStatus { get; set; }
        [DataMember]
        public string WarrantyEndDate { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }

        [DataMember]
        public List<IoTProductWarranty> ProductWarranty { get; set; }

        public List<IoTRegisteredProducts> GetIoTRegisteredProducts(IoTRegisteredProducts registeredProduct, IOrganizationService service)
        {
            IoTRegisteredProducts objRegisteredProducts;
            List<IoTRegisteredProducts> lstRegisteredProducts = new List<IoTRegisteredProducts>();
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                //IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (registeredProduct.CustomerGuid.ToString().Trim().Length == 0)
                    {
                        objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "204", StatusDescription = "Customer GUID is required." };
                        lstRegisteredProducts.Add(objRegisteredProducts);
                        return lstRegisteredProducts;
                    }
                    Query = new QueryExpression("contact");
                    Query.ColumnSet = new ColumnSet("fullname");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("contactid", ConditionOperator.Equal, registeredProduct.CustomerGuid);
                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "204", StatusDescription = "Customer does not exist." };
                        lstRegisteredProducts.Add(objRegisteredProducts);
                    }
                    else
                    {
                        string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_customerasset'>
                            <attribute name='createdon' />
                            <attribute name='msdyn_product' />
                            <attribute name='msdyn_name' />
                            <attribute name='hil_productsubcategorymapping' />
                            <attribute name='hil_productsubcategory' />
                            <attribute name='hil_productcategory' />
                            <attribute name='msdyn_customerassetid' />
                            <attribute name='hil_pincode' />
                            <attribute name='hil_batchnumber' />
                            <attribute name='hil_invoiceavailable' />
                            <attribute name='hil_invoicedate' />
                            <attribute name='hil_invoiceno' />
                            <attribute name='hil_invoicevalue' />
                            <attribute name='hil_purchasedfrom' />
                            <attribute name='hil_retailerpincode' />
                            <attribute name='hil_product' />
                            <attribute name='hil_modelname' />
                            <attribute name='hil_warrantystatus' />
                            <attribute name='hil_warrantysubstatus' />
                            <attribute name='hil_warrantytilldate' />
                            <order attribute='hil_invoicedate' descending='true' />
	                        <filter type='and'>
                                <condition attribute='hil_customer' operator='eq' value='{" + registeredProduct.CustomerGuid.ToString() + @"}' />
                            </filter>
                        </entity>
                        </fetch>";

                        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (entcoll.Entities.Count == 0)
                        {
                            objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "204", StatusDescription = "No Product registered with Customer." };
                            lstRegisteredProducts.Add(objRegisteredProducts);
                        }
                        else
                        {
                            foreach (Entity ent in entcoll.Entities)
                            {
                                objRegisteredProducts = new IoTRegisteredProducts();
                                objRegisteredProducts.CustomerGuid = registeredProduct.CustomerGuid;
                                objRegisteredProducts.RegisteredProductGuid = ent.GetAttributeValue<Guid>("msdyn_customerassetid");

                                if (ent.Attributes.Contains("hil_productcategory"))
                                {
                                    objRegisteredProducts.ProductCategory = ent.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                                    objRegisteredProducts.ProductCategoryId = ent.GetAttributeValue<EntityReference>("hil_productcategory").Id;
                                }
                                if (ent.Attributes.Contains("msdyn_product"))
                                {
                                    objRegisteredProducts.ProductCode = ent.GetAttributeValue<EntityReference>("msdyn_product").Name;
                                    objRegisteredProducts.ProductId = ent.GetAttributeValue<EntityReference>("msdyn_product").Id;
                                }
                                if (ent.Attributes.Contains("hil_modelname"))
                                { objRegisteredProducts.ProductName = ent.GetAttributeValue<string>("hil_modelname"); }
                                if (ent.Attributes.Contains("hil_productsubcategory"))
                                {
                                    objRegisteredProducts.ProductSubCategory = ent.GetAttributeValue<EntityReference>("hil_productsubcategory").Name;
                                    objRegisteredProducts.ProductSubCategoryId = ent.GetAttributeValue<EntityReference>("hil_productsubcategory").Id;
                                }
                                if (ent.Attributes.Contains("msdyn_name"))
                                { objRegisteredProducts.SerialNumber = ent.GetAttributeValue<string>("msdyn_name"); }
                                if (ent.Attributes.Contains("hil_batchnumber"))
                                { objRegisteredProducts.BatchNumber = ent.GetAttributeValue<string>("hil_batchnumber"); }
                                if (ent.Attributes.Contains("hil_invoiceavailable"))
                                { objRegisteredProducts.InvoiceAvailable = ent.GetAttributeValue<bool>("hil_invoiceavailable"); }
                                if (ent.Attributes.Contains("hil_invoiceno"))
                                { objRegisteredProducts.InvoiceNumber = ent.GetAttributeValue<string>("hil_invoiceno"); }
                                if (ent.Attributes.Contains("hil_invoicedate"))
                                { objRegisteredProducts.InvoiceDate = ent.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330).ToShortDateString(); }
                                if (ent.Attributes.Contains("hil_invoicevalue"))
                                { objRegisteredProducts.InvoiceValue = ent.GetAttributeValue<decimal>("hil_invoicevalue"); }
                                if (ent.Attributes.Contains("hil_purchasedfrom"))
                                { objRegisteredProducts.PurchasedFrom = ent.GetAttributeValue<string>("hil_purchasedfrom"); }
                                if (ent.Attributes.Contains("hil_retailerpincode"))
                                { objRegisteredProducts.PurchasedFromLocation = ent.GetAttributeValue<string>("hil_retailerpincode"); }

                                if (ent.Attributes.Contains("hil_product"))
                                { objRegisteredProducts.InstalledLocationEnum = ent.GetAttributeValue<OptionSetValue>("hil_product").Value; }

                                if (ent.Attributes.Contains("hil_product"))
                                { objRegisteredProducts.InstalledLocation = ent.FormattedValues["hil_product"].ToString(); }

                                if (ent.Attributes.Contains("hil_warrantystatus"))
                                {
                                    OptionSetValue _warrantyStatus = ent.GetAttributeValue<OptionSetValue>("hil_warrantystatus");
                                    objRegisteredProducts.WarrantyStatus = _warrantyStatus.Value == 1 ? "In Warranty" : "Out Of Warranty";
                                }
                                if (ent.Attributes.Contains("hil_warrantysubstatus"))
                                {
                                    OptionSetValue _warrantySubStatus = ent.GetAttributeValue<OptionSetValue>("hil_warrantysubstatus");
                                    objRegisteredProducts.WarrantySubStatus = _warrantySubStatus.Value == 1 ? "Standard" : _warrantySubStatus.Value == 2 ? "Extended" : _warrantySubStatus.Value == 3 ? "Special Scheme" : "AMC";
                                }
                                if (ent.Attributes.Contains("hil_warrantytilldate"))
                                {
                                    objRegisteredProducts.WarrantyEndDate = ent.GetAttributeValue<DateTime>("hil_warrantytilldate").AddMinutes(330).ToShortDateString();
                                }

                                #region Getting Product Unit Warranty
                                fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_unitwarranty'>
                                    <attribute name='hil_warrantystartdate' />
                                    <attribute name='hil_warrantyenddate' />
                                    <order attribute='hil_warrantystartdate' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='hil_customerasset' operator='eq' value='" + objRegisteredProducts.RegisteredProductGuid + @"' />
                                      <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                    <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplate' visible='false' link-type='outer' alias='wrt'>
                                      <attribute name='hil_description' />
                                      <attribute name='hil_type' />
                                    </link-entity>
                                  </entity>
                                </fetch>";
                                EntityCollection entcollUW = service.RetrieveMultiple(new FetchExpression(fetchXml));
                                if (entcollUW.Entities.Count > 0)
                                {
                                    objRegisteredProducts.ProductWarranty = new List<IoTProductWarranty>();
                                    OptionSetValue wrtType = new OptionSetValue();
                                    string _wryStartDate = string.Empty;
                                    string _wryEndDate = string.Empty;
                                    string _wrySpecs = string.Empty;
                                    foreach (Entity entUW in entcollUW.Entities)
                                    {
                                        wrtType = (OptionSetValue)entUW.GetAttributeValue<AliasedValue>("wrt.hil_type").Value;
                                        DateTime _wrtDate = entUW.GetAttributeValue<DateTime>("hil_warrantystartdate").AddMinutes(330);
                                        _wryStartDate = _wrtDate.Day.ToString().PadLeft(2, '0') + "/" + _wrtDate.Month.ToString().PadLeft(2, '0') + "/" + _wrtDate.Year.ToString();
                                        _wrtDate = entUW.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                                        _wryEndDate = _wrtDate.Day.ToString().PadLeft(2, '0') + "/" + _wrtDate.Month.ToString().PadLeft(2, '0') + "/" + _wrtDate.Year.ToString();
                                        if (entUW.Attributes.Contains("wrt.hil_description"))
                                        {
                                            _wrySpecs = entUW.GetAttributeValue<AliasedValue>("wrt.hil_description").Value.ToString();
                                        }
                                        objRegisteredProducts.ProductWarranty.Add(
                                            new IoTProductWarranty()
                                            {
                                                WarrantyType = (wrtType.Value == 1 ? "Standard" : wrtType.Value == 2 ? "Extended" : wrtType.Value == 3 ? "AMC" : "Special Scheme"),
                                                WarrantyStartDate = _wryStartDate,
                                                WarrantyEndDate = _wryEndDate,
                                                WarrantySpecifications = _wrySpecs
                                            }
                                            );
                                    }
                                }
                                #endregion
                                objRegisteredProducts.StatusCode = "200";
                                objRegisteredProducts.StatusDescription = "OK";
                                lstRegisteredProducts.Add(objRegisteredProducts);
                            }
                        }
                    }
                }
                else
                {
                    objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    lstRegisteredProducts.Add(objRegisteredProducts);
                }
            }
            catch (Exception ex)
            {
                objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                lstRegisteredProducts.Add(objRegisteredProducts);
            }
            return lstRegisteredProducts;
        }

    }

    [DataContract]
    public class IoTRegisteredProductsResult
    {
        [DataMember]
        public Guid CustomerGuid { get; set; }

        [DataMember]
        public Guid RegisteredProductGuid { get; set; }

        [DataMember]
        public string StatusCode { get; set; }

        [DataMember]
        public string StatusDescription { get; set; }
        [DataMember]
        public string ProductCategory { get; set; }
        [DataMember]
        public Guid ProductCategoryId { get; set; }
        [DataMember]
        public string ProductSubCategory { get; set; }
        [DataMember]
        public Guid ProductSubCategoryId { get; set; }
    }

    [DataContract]
    public class IoTProductWarranty {
        [DataMember]
        public string WarrantyType { get; set; }
        [DataMember]
        public string WarrantyStartDate { get; set; }
        [DataMember]
        public string WarrantyEndDate { get; set; }
        [DataMember]
        public string WarrantySpecifications { get; set; }
    }

    [DataContract]
    public class SerialNumberData
    {
        public string SerialNumber { get; set; }
    }
}
