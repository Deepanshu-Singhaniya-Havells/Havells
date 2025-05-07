using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Customer_Asset
{
    public class RegisteredProducts : IPlugin
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
            List<IoTRegisteredProducts> lstRegisteredProducts = new List<IoTRegisteredProducts>();
            Guid CustomerGuid = Guid.Empty;

            if (context.InputParameters.Contains("CustomerGuid") && context.InputParameters["CustomerGuid"] is string)
            {
                bool isValidCustomerGuid = Guid.TryParse(context.InputParameters["CustomerGuid"].ToString(), out CustomerGuid);
                if (!isValidCustomerGuid)
                {
                    lstRegisteredProducts.Add(new IoTRegisteredProducts
                    {
                        StatusCode = "204",
                        StatusDescription = "Invalid Customer GUID."
                    });
                    JsonResponse = JsonSerializer.Serialize(lstRegisteredProducts);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                JsonResponse = JsonSerializer.Serialize(GetIoTRegisteredProducts(service, CustomerGuid));
                context.OutputParameters["data"] = JsonResponse;
            }
            else
            {
                lstRegisteredProducts.Add(new IoTRegisteredProducts
                {
                    StatusCode = "204",
                    StatusDescription = "Customer GUID is required."
                });
                JsonResponse = JsonSerializer.Serialize(lstRegisteredProducts);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
        }
        public List<IoTRegisteredProducts> GetIoTRegisteredProducts(IOrganizationService service, Guid CustomerGuid)
        {
            IoTRegisteredProducts objRegisteredProducts;
            List<IoTRegisteredProducts> lstRegisteredProducts = new List<IoTRegisteredProducts>();
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                if (service != null)
                {
                    Query = new QueryExpression("msdyn_customerasset");
                    Query.ColumnSet = new ColumnSet("hil_warrantytilldate", "hil_warrantysubstatus", "hil_warrantystatus", "hil_modelname", "hil_product", "hil_retailerpincode", "hil_purchasedfrom", "hil_invoicevalue", "hil_invoiceno", "hil_invoicedate", "hil_invoiceavailable", "hil_batchnumber", "hil_pincode", "msdyn_customerassetid", "createdon", "msdyn_product", "msdyn_name", "hil_productsubcategorymapping", "hil_productsubcategory", "hil_productcategory");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, CustomerGuid);
                    Query.AddOrder("hil_invoicedate", OrderType.Descending);
                    Query.TopCount = 30;

                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objRegisteredProducts = new IoTRegisteredProducts { StatusCode = "204", StatusDescription = "Customer Product does not exist." };
                        lstRegisteredProducts.Add(objRegisteredProducts);
                    }
                    else
                    {
                        foreach (Entity ent in entcoll.Entities)
                        {
                            objRegisteredProducts = new IoTRegisteredProducts();
                            objRegisteredProducts.DealerPinCode = "";
                            objRegisteredProducts.BatchNumber = "";
                            objRegisteredProducts.CustomerGuid = CustomerGuid;
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
    public class IoTRegisteredProducts
    {

        public Guid CustomerGuid { get; set; }

        public Guid RegisteredProductGuid { get; set; }

        public string ProductCategory { get; set; }

        public Guid ProductCategoryId { get; set; }

        public string ProductSubCategory { get; set; }

        public Guid ProductSubCategoryId { get; set; }

        public string DealerPinCode { get; set; }

        public Guid ProductId { get; set; }

        public string ProductCode { get; set; }

        public string ProductName { get; set; }

        public string SerialNumber { get; set; }

        public string BatchNumber { get; set; }

        public bool InvoiceAvailable { get; set; }

        public string InvoiceNumber { get; set; }

        public string InvoiceDate { get; set; }

        public Decimal? InvoiceValue { get; set; }

        public string PurchasedFrom { get; set; }

        public string PurchasedFromLocation { get; set; }

        public string InstalledLocation { get; set; }

        public int InstalledLocationEnum { get; set; }

        public int SourceOfRegistration { get; set; }

        public string WarrantyStatus { get; set; }

        public string WarrantySubStatus { get; set; }

        public string WarrantyEndDate { get; set; }

        public string StatusCode { get; set; }

        public string StatusDescription { get; set; }


        public List<IoTProductWarranty> ProductWarranty { get; set; }
    }
    public class IoTProductWarranty
    {
        public string WarrantyType { get; set; }
        public string WarrantyStartDate { get; set; }
        public string WarrantyEndDate { get; set; }
        public string WarrantySpecifications { get; set; }
    }

}
