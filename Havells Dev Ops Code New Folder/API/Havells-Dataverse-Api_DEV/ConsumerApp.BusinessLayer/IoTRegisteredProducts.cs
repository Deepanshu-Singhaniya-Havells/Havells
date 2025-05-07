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

namespace ConsumerApp.BusinessLayer
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
        //public List<IoTRegisteredProducts> GetIoTRegisteredProducts(IoTRegisteredProducts registeredProduct)
        //{
        //    IoTRegisteredProducts objRegisteredProducts;
        //    List<IoTRegisteredProducts> lstRegisteredProducts = new List<IoTRegisteredProducts>();
        //    EntityCollection entcoll;
        //    QueryExpression Query;
        //    try
        //    {
        //        IOrganizationService service = ConnectToCRM.GetOrgService();
        //        if (service != null)
        //        {
        //            if (registeredProduct.CustomerGuid.ToString().Trim().Length == 0)
        //            {
        //                objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "204", StatusDescription = "Customer GUID is required." };
        //                lstRegisteredProducts.Add(objRegisteredProducts);
        //                return lstRegisteredProducts;
        //            }
        //            Query = new QueryExpression("msdyn_customerasset");
        //            Query.ColumnSet = new ColumnSet("hil_warrantytilldate", "hil_warrantysubstatus", "hil_warrantystatus", "hil_modelname", "hil_product", "hil_retailerpincode", "hil_purchasedfrom", "hil_invoicevalue", "hil_invoiceno", "hil_invoicedate", "hil_invoiceavailable", "hil_batchnumber", "createdon", "msdyn_product", "msdyn_name", "hil_productsubcategory", "hil_productcategory");
        //            ////***changed by Saurabh
        //            //Query.ColumnSet = new ColumnSet("hil_warrantytilldate", "hil_warrantysubstatus", "hil_warrantystatus", "hil_modelname", "hil_product", "hil_retailerpincode", "hil_purchasedfrom", "hil_invoicevalue", "hil_invoiceno", "hil_invoicedate", "hil_invoiceavailable", "hil_batchnumber", "hil_pincode", "msdyn_customerassetid", "createdon", "msdyn_product", "msdyn_name", "hil_productsubcategorymapping", "hil_productsubcategory", "hil_productcategory");
        //            //***changed by Saurabh ends here.
        //            Query.Criteria = new FilterExpression(LogicalOperator.And);
        //            Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, registeredProduct.CustomerGuid);
        //            Query.AddOrder("createdon", OrderType.Descending);
        //            Query.TopCount = 10;

        //            entcoll = service.RetrieveMultiple(Query);
        //            if (entcoll.Entities.Count == 0)
        //            {
        //                objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "204", StatusDescription = "Customer does not exist." };
        //                lstRegisteredProducts.Add(objRegisteredProducts);
        //            }
        //            else
        //            {
        //                foreach (Entity ent in entcoll.Entities)
        //                {
        //                    objRegisteredProducts = new IoTRegisteredProducts();
        //                    objRegisteredProducts.CustomerGuid = registeredProduct.CustomerGuid;
        //                    //***changed by Saurabh
        //                    //objRegisteredProducts.RegisteredProductGuid = ent.GetAttributeValue<Guid>("msdyn_customerassetid");
        //                    objRegisteredProducts.RegisteredProductGuid = ent.Id;
        //                    //***changed by Saurabh ends here.
        //                    if (ent.Attributes.Contains("hil_productcategory"))
        //                    {
        //                        objRegisteredProducts.ProductCategory = ent.GetAttributeValue<EntityReference>("hil_productcategory").Name;
        //                        objRegisteredProducts.ProductCategoryId = ent.GetAttributeValue<EntityReference>("hil_productcategory").Id;
        //                    }
        //                    if (ent.Attributes.Contains("msdyn_product"))
        //                    {
        //                        objRegisteredProducts.ProductCode = ent.GetAttributeValue<EntityReference>("msdyn_product").Name;
        //                        objRegisteredProducts.ProductId = ent.GetAttributeValue<EntityReference>("msdyn_product").Id;
        //                    }
        //                    if (ent.Attributes.Contains("hil_modelname"))
        //                    { objRegisteredProducts.ProductName = ent.GetAttributeValue<string>("hil_modelname"); }
        //                    if (ent.Attributes.Contains("hil_productsubcategory"))
        //                    {
        //                        objRegisteredProducts.ProductSubCategory = ent.GetAttributeValue<EntityReference>("hil_productsubcategory").Name;
        //                        objRegisteredProducts.ProductSubCategoryId = ent.GetAttributeValue<EntityReference>("hil_productsubcategory").Id;
        //                    }
        //                    if (ent.Attributes.Contains("msdyn_name"))
        //                    { objRegisteredProducts.SerialNumber = ent.GetAttributeValue<string>("msdyn_name"); }
        //                    if (ent.Attributes.Contains("hil_batchnumber"))
        //                    { objRegisteredProducts.BatchNumber = ent.GetAttributeValue<string>("hil_batchnumber"); }
        //                    if (ent.Attributes.Contains("hil_invoiceavailable"))
        //                    { objRegisteredProducts.InvoiceAvailable = ent.GetAttributeValue<bool>("hil_invoiceavailable"); }
        //                    if (ent.Attributes.Contains("hil_invoiceno"))
        //                    { objRegisteredProducts.InvoiceNumber = ent.GetAttributeValue<string>("hil_invoiceno"); }
        //                    if (ent.Attributes.Contains("hil_invoicedate"))
        //                    { objRegisteredProducts.InvoiceDate = ent.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330).ToShortDateString(); }
        //                    if (ent.Attributes.Contains("hil_invoicevalue"))
        //                    { objRegisteredProducts.InvoiceValue = ent.GetAttributeValue<decimal>("hil_invoicevalue"); }
        //                    if (ent.Attributes.Contains("hil_purchasedfrom"))
        //                    { objRegisteredProducts.PurchasedFrom = ent.GetAttributeValue<string>("hil_purchasedfrom"); }
        //                    if (ent.Attributes.Contains("hil_retailerpincode"))
        //                    { objRegisteredProducts.PurchasedFromLocation = ent.GetAttributeValue<string>("hil_retailerpincode"); }

        //                    if (ent.Attributes.Contains("hil_product"))
        //                    { 
        //                        objRegisteredProducts.InstalledLocationEnum = ent.GetAttributeValue<OptionSetValue>("hil_product").Value;
        //                        objRegisteredProducts.InstalledLocation = ent.FormattedValues["hil_product"];
        //                    }

        //                    //***changed by Saurabh
        //                    //if (ent.Attributes.Contains("hil_product"))
        //                    //{ objRegisteredProducts.InstalledLocationEnum = ent.GetAttributeValue<OptionSetValue>("hil_product").Value; }

        //                    //if (ent.Attributes.Contains("hil_product"))
        //                    //{ objRegisteredProducts.InstalledLocation = ent.FormattedValues["hil_product"]; }
        //                    //***changed by Saurabh ends here.

        //                    if (ent.Attributes.Contains("hil_warrantystatus"))
        //                    {
        //                        OptionSetValue _warrantyStatus = ent.GetAttributeValue<OptionSetValue>("hil_warrantystatus");
        //                        objRegisteredProducts.WarrantyStatus = _warrantyStatus.Value == 1 ? "In Warranty" : "Out Warranty";
        //                    }
        //                    if (ent.Attributes.Contains("hil_warrantysubstatus"))
        //                    {
        //                        OptionSetValue _warrantySubStatus = ent.GetAttributeValue<OptionSetValue>("hil_warrantysubstatus");
        //                        objRegisteredProducts.WarrantySubStatus = _warrantySubStatus.Value == 1 ? "Standard" : _warrantySubStatus.Value == 2 ? "Extended" : _warrantySubStatus.Value == 3 ? "Special Scheme" : "AMC";
        //                    }
        //                    if (ent.Attributes.Contains("hil_warrantytilldate"))
        //                    {
        //                        objRegisteredProducts.WarrantyEndDate = ent.GetAttributeValue<DateTime>("hil_warrantytilldate").AddMinutes(330).ToShortDateString();
        //                    }
        //                    objRegisteredProducts.StatusCode = "200";
        //                    objRegisteredProducts.StatusDescription = "OK";
        //                    lstRegisteredProducts.Add(objRegisteredProducts);
        //                }
        //            }
        //        }
        //        else
        //        {
        //            objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
        //            lstRegisteredProducts.Add(objRegisteredProducts);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
        //        lstRegisteredProducts.Add(objRegisteredProducts);
        //    }
        //    return lstRegisteredProducts;
        //}

        public List<IoTRegisteredProducts> GetIoTRegisteredProducts(IoTRegisteredProducts registeredProduct)
        {
            IoTRegisteredProducts objRegisteredProducts;
            List<IoTRegisteredProducts> lstRegisteredProducts = new List<IoTRegisteredProducts>();
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)service).IsReady)
                {
                    if (registeredProduct.CustomerGuid.ToString().Trim().Length == 0)
                    {
                        objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "204", StatusDescription = "Customer GUID is required." };
                        lstRegisteredProducts.Add(objRegisteredProducts);
                        return lstRegisteredProducts;
                    }

                    //Query = new QueryExpression("contact");
                    //Query.ColumnSet = new ColumnSet("fullname");
                    //Query.Criteria = new FilterExpression(LogicalOperator.And);
                    //Query.Criteria.AddCondition("contactid", ConditionOperator.Equal, registeredProduct.CustomerGuid);
                    Query = new QueryExpression("msdyn_customerasset");
                    Query.ColumnSet = new ColumnSet("hil_warrantytilldate", "hil_warrantysubstatus", "hil_warrantystatus", "hil_modelname", "hil_product", "hil_retailerpincode", "hil_purchasedfrom", "hil_invoicevalue", "hil_invoiceno", "hil_invoicedate", "hil_invoiceavailable", "hil_batchnumber", "hil_pincode", "msdyn_customerassetid", "createdon", "msdyn_product", "msdyn_name", "hil_productsubcategorymapping", "hil_productsubcategory", "hil_productcategory");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, registeredProduct.CustomerGuid);
                    Query.AddOrder("hil_invoicedate", OrderType.Descending);
                    Query.TopCount = 20;

                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "204", StatusDescription = "Customer Product does not exist." };
                        lstRegisteredProducts.Add(objRegisteredProducts);
                    }
                    else
                    {
                        //string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        //<entity name='msdyn_customerasset'>
                        //    <attribute name='createdon' />
                        //    <attribute name='msdyn_product' />
                        //    <attribute name='msdyn_name' />
                        //    <attribute name='hil_productsubcategorymapping' />
                        //    <attribute name='hil_productsubcategory' />
                        //    <attribute name='hil_productcategory' />
                        //    <attribute name='msdyn_customerassetid' />
                        //    <attribute name='hil_pincode' />
                        //    <attribute name='hil_batchnumber' />
                        //    <attribute name='hil_invoiceavailable' />
                        //    <attribute name='hil_invoicedate' />
                        //    <attribute name='hil_invoiceno' />
                        //    <attribute name='hil_invoicevalue' />
                        //    <attribute name='hil_purchasedfrom' />
                        //    <attribute name='hil_retailerpincode' />
                        //    <attribute name='hil_product' />
                        //    <attribute name='hil_modelname' />
                        //    <attribute name='hil_warrantystatus' />
                        //    <attribute name='hil_warrantysubstatus' />
                        //    <attribute name='hil_warrantytilldate' />
                        //    <order attribute='hil_invoicedate' descending='true' />
                        // <filter type='and'>
                        //        <condition attribute='hil_customer' operator='eq' value='{" + registeredProduct.CustomerGuid.ToString() + @"}' />
                        //    </filter>
                        //</entity>
                        //</fetch>";

                        //entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        //if (entcoll.Entities.Count == 0)
                        //{
                        //    objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "204", StatusDescription = "No Product registered with Customer." };
                        //    lstRegisteredProducts.Add(objRegisteredProducts);
                        //}
                        //else
                        //{
                        foreach (Entity ent in entcoll.Entities)
                        {
                            objRegisteredProducts = new IoTRegisteredProducts();
                            objRegisteredProducts.DealerPinCode = "";
                            objRegisteredProducts.BatchNumber = "";
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
                            { objRegisteredProducts.InvoiceValue = Math.Round(ent.GetAttributeValue<decimal>("hil_invoicevalue"), 2); }
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
                            else
                            {
                                objRegisteredProducts.WarrantyStatus = "Pending for Approval";
                                objRegisteredProducts.WarrantySubStatus = "";
                                objRegisteredProducts.WarrantyEndDate = "";
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
                            string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
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
                            objRegisteredProducts.ProductWarranty = new List<IoTProductWarranty>();
                            if (entcollUW.Entities.Count > 0)
                            {
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
                        //}
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

        public List<IoTRegisteredProducts> GetIoTRegisteredProductsProd(IoTRegisteredProducts registeredProduct)
        {
            IoTRegisteredProducts objRegisteredProducts;
            List<IoTRegisteredProducts> lstRegisteredProducts = new List<IoTRegisteredProducts>();
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                IOrganizationService service = ConnectToCRM.GetCRMConnection();
                if (service != null)
                {
                    if (registeredProduct.CustomerGuid.ToString().Trim().Length == 0)
                    {
                        objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "204", StatusDescription = "Customer GUID is required." };
                        lstRegisteredProducts.Add(objRegisteredProducts);
                        return lstRegisteredProducts;
                    }

                    Query = new QueryExpression("msdyn_customerasset");
                    Query.ColumnSet = new ColumnSet("hil_warrantytilldate", "hil_warrantysubstatus", "hil_warrantystatus", "hil_modelname", "hil_product", "hil_retailerpincode", "hil_purchasedfrom", "hil_invoicevalue", "hil_invoiceno", "hil_invoicedate", "hil_invoiceavailable", "hil_batchnumber", "hil_pincode", "msdyn_customerassetid", "createdon", "msdyn_product", "msdyn_name", "hil_productsubcategorymapping", "hil_productsubcategory", "hil_productcategory");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, registeredProduct.CustomerGuid);
                    Query.AddOrder("createdon", OrderType.Descending);
                    Query.TopCount = 10;

                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "204", StatusDescription = "Customer or  does not exist." };
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
                                objRegisteredProducts.WarrantyStatus = _warrantyStatus.Value == 1 ? "In Warranty" : "Out Warranty";
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
                            objRegisteredProducts.StatusCode = "200";
                            objRegisteredProducts.StatusDescription = "OK";
                            lstRegisteredProducts.Add(objRegisteredProducts);
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

        public IoTRegisteredProductsResult RegisterProduct(IoTRegisteredProducts productData)
        {
            IoTRegisteredProductsResult objRegisteredProductResult = null;
            QueryExpression Query;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (productData.CustomerGuid == Guid.Empty)
                    {
                        objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Customer does not exist." };
                        return objRegisteredProductResult;
                    }
                    if (productData.SerialNumber != null && productData.SerialNumber.Trim().Length > 0)
                    {
                        Entity entAsset = CheckIfExistingSerialNumberWithDetails(service, productData.SerialNumber);
                        if (entAsset != null)
                        {
                            objRegisteredProductResult = new IoTRegisteredProductsResult { RegisteredProductGuid = entAsset.Id, ProductCategory = entAsset.GetAttributeValue<EntityReference>("hil_productcategory").Name, ProductCategoryId = entAsset.GetAttributeValue<EntityReference>("hil_productcategory").Id, ProductSubCategory = entAsset.GetAttributeValue<EntityReference>("hil_productsubcategory").Name, ProductSubCategoryId = entAsset.GetAttributeValue<EntityReference>("hil_productsubcategory").Id, CustomerGuid = productData.CustomerGuid, StatusCode = "208", StatusDescription = "SERIAL NUMBER ALREADY REGISTERED." };
                            return objRegisteredProductResult;
                        }
                    }
                    if (productData.ProductCategoryId == Guid.Empty)
                    {
                        objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Product Category is required." };
                        return objRegisteredProductResult;
                    }
                    if (productData.ProductSubCategoryId == Guid.Empty)
                    {
                        objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Product Subcategory is required." };
                        return objRegisteredProductResult;
                    }
                    if (productData.ProductId == Guid.Empty)
                    {
                        objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Product is required." };
                        return objRegisteredProductResult;
                    }
                    if (productData.InvoiceAvailable)
                    {
                        if (productData.InvoiceNumber == null || productData.InvoiceNumber.Trim().Length == 0)
                        {
                            objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Invoice No. is required." };
                            return objRegisteredProductResult;
                        }
                        if (productData.InvoiceDate == null)
                        {
                            objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Invoice Date is required." };
                            return objRegisteredProductResult;
                        }
                        if (productData.InvoiceValue == null || productData.InvoiceValue == 0)
                        {
                            objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Invoice Value is required." };
                            return objRegisteredProductResult;
                        }
                        if (productData.PurchasedFrom == null || productData.PurchasedFrom.Trim().Length == 0)
                        {
                            objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Purchased From is required." };
                            return objRegisteredProductResult;
                        }
                        if (productData.PurchasedFromLocation == null || productData.PurchasedFromLocation.Trim().Length == 0)
                        {
                            objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Purchased From Location(PIN Code) is required." };
                            return objRegisteredProductResult;
                        }
                    }
                    //bool IfExisting = CheckIfExistingSerialNumber(service, productData.SerialNumber);
                    //if (IfExisting == false)
                    //{
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        Guid registeredProductId = Guid.Empty;

                        msdyn_customerasset entCustomerAsset = new msdyn_customerasset();
                        entCustomerAsset.hil_Customer = new EntityReference("contact", productData.CustomerGuid);
                        if (productData.SerialNumber != null)
                        {
                            entCustomerAsset.msdyn_name = productData.SerialNumber;
                        }
                        entCustomerAsset["hil_source"] = new OptionSetValue(4); //IoT Platform
                        entCustomerAsset.hil_InvoiceAvailable = productData.InvoiceAvailable;
                        entCustomerAsset["hil_retailerpincode"] = productData.PurchasedFromLocation;
                        entCustomerAsset["hil_purchasedfrom"] = productData.PurchasedFrom;
                        entCustomerAsset.msdyn_Product = new EntityReference("product", productData.ProductId);
                        entCustomerAsset.hil_ModelName = productData.ProductName;
                        entCustomerAsset.hil_ProductCategory = new EntityReference("product", productData.ProductCategoryId);
                        entCustomerAsset.hil_ProductSubcategory = new EntityReference("product", productData.ProductSubCategoryId);

                        Query = new QueryExpression(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName);
                        Query.ColumnSet = new ColumnSet("hil_productcategorydivision", "hil_name", "hil_productsubcategorymg", "statecode");
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition("hil_productcategorydivision", ConditionOperator.Equal, productData.ProductCategoryId);
                        Query.Criteria.AddCondition("hil_productsubcategorymg", ConditionOperator.Equal, productData.ProductSubCategoryId);
                        EntityCollection Found = service.RetrieveMultiple(Query);
                        if (Found.Entities.Count > 0)
                        {
                            entCustomerAsset.hil_productsubcategorymapping = Found.Entities[0].ToEntityReference();
                        }
                        entCustomerAsset.statuscode = new OptionSetValue(910590000); // Pending for Approval
                        if (productData.InvoiceAvailable)
                        {
                            if (productData.InvoiceDate != null)
                            {
                                DateTime dtInvoice = DateTime.ParseExact(productData.InvoiceDate, new string[] { "MM.dd.yyyy", "MM-dd-yyyy", "MM/dd/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None);
                                entCustomerAsset.hil_InvoiceDate = dtInvoice;
                            }
                            if (productData.InvoiceNumber != string.Empty)
                            {
                                entCustomerAsset.hil_InvoiceNo = productData.InvoiceNumber;
                            }
                            if (productData.InvoiceValue != null)
                            {
                                entCustomerAsset.hil_InvoiceValue = productData.InvoiceValue;
                            }
                        }
                        registeredProductId = service.Create(entCustomerAsset);
                        if (registeredProductId != Guid.Empty)
                        {
                            objRegisteredProductResult = new IoTRegisteredProductsResult { ProductCategory = entCustomerAsset.hil_ProductCategory.Name, ProductSubCategory = entCustomerAsset.hil_ProductSubcategory.Name, ProductCategoryId = productData.ProductCategoryId, ProductSubCategoryId = productData.ProductSubCategoryId, CustomerGuid = productData.CustomerGuid, RegisteredProductGuid = registeredProductId, StatusCode = "200", StatusDescription = "OK" };
                        }
                    }
                }
                else
                {
                    objRegisteredProductResult = new IoTRegisteredProductsResult { CustomerGuid = productData.CustomerGuid, StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                objRegisteredProductResult = new IoTRegisteredProductsResult { CustomerGuid = productData.CustomerGuid, StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return objRegisteredProductResult;
        }

        public IoTRegisteredProductsResult RegisterProductFromWhatsapp(IoTRegisteredProducts productData)
        {
            IoTRegisteredProductsResult objRegisteredProductResult = null;
            EntityCollection entcoll;
            QueryExpression Query;
            EntityReference productCategory = null;
            EntityReference productSubCategory = null;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (productData.CustomerGuid == Guid.Empty)
                    {
                        objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Customer does not exist." };
                        return objRegisteredProductResult;
                    }
                    if (productData.SerialNumber == null || productData.SerialNumber.Trim().Length == 0)
                    {
                        objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Serial Number is required." };
                        return objRegisteredProductResult;
                    }
                    if (productData.ProductCode == null || productData.ProductCode.Trim().Length == 0)
                    {
                        productData.ProductCode = "OTHER";
                    }
                    if (productData.InvoiceAvailable)
                    {
                        if (productData.InvoiceDate == null)
                        {
                            objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Invoice Date is required." };
                            return objRegisteredProductResult;
                        }
                        if (productData.SourceOfRegistration != 5 && productData.SourceOfRegistration != 7) // Applicable for Other than Whatsapp and ChatBot Source
                        {
                            if (productData.InvoiceNumber == null || productData.InvoiceNumber.Trim().Length == 0)
                            {
                                objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Invoice No. is required." };
                                return objRegisteredProductResult;
                            }
                            if (productData.InvoiceValue == null || productData.InvoiceValue == 0)
                            {
                                objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Invoice Value is required." };
                                return objRegisteredProductResult;
                            }
                            if (productData.PurchasedFrom == null || productData.PurchasedFrom.Trim().Length == 0)
                            {
                                objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Purchased From is required." };
                                return objRegisteredProductResult;
                            }
                            if (productData.PurchasedFromLocation == null || productData.PurchasedFromLocation.Trim().Length == 0)
                            {
                                objRegisteredProductResult = new IoTRegisteredProductsResult { StatusCode = "204", StatusDescription = "Purchased From Location(PIN Code) is required." };
                                return objRegisteredProductResult;
                            }
                        }
                    }
                    Entity entAsset = CheckIfExistingSerialNumberWithDetails(service, productData.SerialNumber);
                    if (entAsset == null)
                    {
                        using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                        {
                            Guid registeredProductId = Guid.Empty;

                            msdyn_customerasset entCustomerAsset = new msdyn_customerasset();
                            entCustomerAsset.hil_Customer = new EntityReference("contact", productData.CustomerGuid);
                            entCustomerAsset.msdyn_name = productData.SerialNumber;
                            entCustomerAsset["hil_source"] = new OptionSetValue(productData.SourceOfRegistration); //WhatsApp For Service
                            entCustomerAsset.hil_InvoiceAvailable = productData.InvoiceAvailable;
                            entCustomerAsset["hil_retailerpincode"] = "0";
                            entCustomerAsset["hil_purchasedfrom"] = "Registered From Whatsapp/OneApp";

                            if (productData.InvoiceAvailable)
                            {
                                if (productData.InvoiceDate != null)
                                {
                                    string _date = productData.InvoiceDate;

                                    DateTime dtInvoice = new DateTime(Convert.ToInt32(_date.Substring(6, 4)), Convert.ToInt32(_date.Substring(0, 2)), Convert.ToInt32(_date.Substring(3, 2)));
                                    //DateTime dtInvoice = DateTime.ParseExact(productData.InvoiceDate, new string[] { "MM.dd.yyyy", "MM-dd-yyyy", "MM/dd/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None);
                                    entCustomerAsset.hil_InvoiceDate = dtInvoice;
                                }
                                if (productData.InvoiceNumber != string.Empty)
                                {
                                    entCustomerAsset.hil_InvoiceNo = productData.InvoiceNumber;
                                }
                                if (productData.InvoiceValue != null)
                                {
                                    entCustomerAsset.hil_InvoiceValue = productData.InvoiceValue;
                                }
                                if (productData.PurchasedFrom != null)
                                {
                                    entCustomerAsset["hil_purchasedfrom"] = productData.PurchasedFrom;
                                }
                                if (productData.PurchasedFromLocation != null)
                                {
                                    entCustomerAsset["hil_retailerpincode"] = productData.PurchasedFromLocation;
                                }
                            }
                            if (productData.InstalledLocationEnum > 0)
                            {
                                entCustomerAsset["hil_product"] = new OptionSetValue(productData.InstalledLocationEnum);
                            }

                            #region UpdateMaterialCodeonAsset
                            var obj1 = (from _Product in orgContext.CreateQuery<Product>()
                                        where _Product.ProductNumber == productData.ProductCode
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
                                entCustomerAsset.msdyn_Product = new EntityReference("product", iobj1.ProductId.Value);
                                entCustomerAsset.hil_ModelName = iobj1.Description;
                                entCustomerAsset.hil_ProductCategory = iobj1.hil_Division;
                                entCustomerAsset.hil_ProductSubcategory = iobj1.hil_MaterialGroup;
                                productCategory = iobj1.hil_Division;
                                productSubCategory = iobj1.hil_MaterialGroup;

                                Query = new QueryExpression(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName);
                                Query.ColumnSet = new ColumnSet("hil_productcategorydivision", "hil_name", "hil_productsubcategorymg", "statecode");
                                Query.Criteria = new FilterExpression(LogicalOperator.And);
                                Query.Criteria.AddCondition("hil_productcategorydivision", ConditionOperator.Equal, iobj1.hil_Division.Id);
                                Query.Criteria.AddCondition("hil_productsubcategorymg", ConditionOperator.Equal, iobj1.hil_MaterialGroup.Id);
                                EntityCollection Found = service.RetrieveMultiple(Query);
                                if (Found.Entities.Count > 0)
                                {
                                    entCustomerAsset.hil_productsubcategorymapping = Found.Entities[0].ToEntityReference();
                                }
                            }
                            #endregion
                            entCustomerAsset.statuscode = new OptionSetValue(910590000); // Pending for Approval
                            registeredProductId = service.Create(entCustomerAsset);

                            if (productCategory != null)
                            {
                                objRegisteredProductResult = new IoTRegisteredProductsResult { ProductCategory = productCategory.Name, ProductCategoryId = productCategory.Id, ProductSubCategory = productSubCategory.Name, ProductSubCategoryId = productSubCategory.Id, CustomerGuid = productData.CustomerGuid, RegisteredProductGuid = registeredProductId, StatusCode = "200", StatusDescription = "OK" };
                            }
                            else
                            {
                                objRegisteredProductResult = new IoTRegisteredProductsResult { CustomerGuid = productData.CustomerGuid, RegisteredProductGuid = registeredProductId, StatusCode = "200", StatusDescription = "OK" };
                            }
                        }
                    }
                    else
                    {
                        //objRegisteredProductResult = new IoTRegisteredProductsResult {CustomerGuid = productData.CustomerGuid, StatusCode = "208", StatusDescription = "SERIAL NUMBER ALREADY REGISTERED." };
                        objRegisteredProductResult = new IoTRegisteredProductsResult
                        {
                            RegisteredProductGuid = entAsset.Id,
                            ProductCategory = entAsset.Contains("hil_productcategory") ? entAsset.GetAttributeValue<EntityReference>("hil_productcategory").Name : String.Empty,
                            ProductCategoryId = entAsset.Contains("hil_productcategory") ? entAsset.GetAttributeValue<EntityReference>("hil_productcategory").Id : Guid.Empty,
                            ProductSubCategory = entAsset.Contains("hil_productsubcategory") ? entAsset.GetAttributeValue<EntityReference>("hil_productsubcategory").Name : String.Empty,
                            ProductSubCategoryId = entAsset.Contains("hil_productsubcategory") ? entAsset.GetAttributeValue<EntityReference>("hil_productsubcategory").Id : Guid.Empty,
                            CustomerGuid = productData.CustomerGuid,
                            StatusCode = "208",
                            StatusDescription = "SERIAL NUMBER ALREADY REGISTERED."
                        };
                    }
                }
                else
                {
                    objRegisteredProductResult = new IoTRegisteredProductsResult { CustomerGuid = productData.CustomerGuid, StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                objRegisteredProductResult = new IoTRegisteredProductsResult { CustomerGuid = productData.CustomerGuid, StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return objRegisteredProductResult;
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
            Query.ColumnSet = new ColumnSet("hil_productcategory", "hil_productsubcategory");
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
    public class IoTProductWarranty
    {
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
