using Microsoft.Xrm.Sdk.Query;
using HavellsSync_ModelData.Common;
using HavellsSync_Data.IManager;
using HavellsSync_ModelData.Product;
using System.Drawing.Text;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using Newtonsoft.Json;
using System.Globalization;
using System.Runtime.Serialization;
using System.Runtime.Intrinsics.Arm;
using Microsoft.VisualBasic;
using System.Collections.Generic;
using Microsoft.VisualBasic.FileIO;
using System.Drawing.Drawing2D;
using System.Text.RegularExpressions;

namespace HavellsSync_Data.Manager
{
    public class ProductManager : IProductManager
    {
        private readonly ICrmService _crmService;

        public ProductManager(ICrmService crmService)
        {
            Check.Argument.IsNotNull(nameof(crmService), crmService);

            this._crmService = crmService;
        }

        public async Task<ValidateSerialNumResponse> ValidateSerialNumber(string SerialNumber, string SessionId)
        {
            ValidateSerialNumResponse result;
            try
            {
                if (_crmService != null)
                {
                    string expiryTime = CommonMethods.getExpiryTime(_crmService, SessionId).ToString();
                    if (!string.IsNullOrWhiteSpace(SerialNumber))
                    {
                        Regex re = new Regex("^([0-9a-zA-Z]){1,32}$");
                        if (!re.IsMatch(SerialNumber))
                        {
                            return new ValidateSerialNumResponse
                            {
                                TokenExpiresAt = expiryTime,
                                StatusCode = 204,
                                Message = CommonMessage.InvalidSerialnumMsg,
                                //Message = "Invalid Serial Number (IDU)."
                            };
                        }
                    }
                    if (string.IsNullOrEmpty(SerialNumber))
                    {
                        return new ValidateSerialNumResponse
                        {
                            TokenExpiresAt = expiryTime,
                            StatusCode = 204,
                            Message = CommonMessage.SerialnumberMsg,
                            //Message = "Asset serial number is required."
                        };
                    }
                    Entity entity = CheckIfExistingSerialNumberWithDetails(_crmService, SerialNumber);
                    if (entity == null)
                    {
                        string? Username = string.Empty;
                        string? Password = string.Empty;
                        string? Url = string.Empty;
                        QueryExpression query = new QueryExpression("hil_integrationconfiguration");
                        query.ColumnSet = new ColumnSet("hil_username", "hil_password", "hil_url");
                        //query.TopCount = 1;
                        query.Criteria = new FilterExpression(LogicalOperator.Or);
                        query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, "Credentials"));
                        query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, "SerialNumberValidation"));
                        query.AddOrder("hil_url", OrderType.Ascending);
                        EntityCollection EntityColl = _crmService.RetrieveMultiple(query);
                        if (EntityColl.Entities.Count > 0)
                        {
                            Username = EntityColl.Entities[0].Contains("hil_username") ? EntityColl.Entities[0].GetAttributeValue<string>("hil_username") : null;
                            Password = EntityColl.Entities[0].Contains("hil_password") ? EntityColl.Entities[0].GetAttributeValue<string>("hil_password") : null;
                            Url = EntityColl.Entities[1].Contains("hil_url") ? EntityColl.Entities[1].GetAttributeValue<string>("hil_url") : null;
                        }

                        if (Url != null)
                        {
                            string address = Url + SerialNumber;
                            WebClient webClient = new WebClient();
                            string text3 = Convert.ToBase64String(Encoding.ASCII.GetBytes(Username + ":" + Password));
                            webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + text3;
                            webClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                            byte[] buffer = webClient.DownloadData(address);
                            DataContractJsonSerializer dataContractJsonSerializer = new DataContractJsonSerializer(typeof(SerialNumberValidation));
                            SerialNumberValidation rootObject = (SerialNumberValidation)dataContractJsonSerializer.ReadObject(new MemoryStream(buffer));
                            if (rootObject.EX_PRD_DET != null)
                            {
                                query = new QueryExpression("product");
                                query.ColumnSet = new ColumnSet("hil_materialgroup", "hil_division", "productnumber", "description", "hil_brandidentifier");
                                query.TopCount = 1;
                                query.Criteria.AddCondition(new ConditionExpression("productnumber", ConditionOperator.Equal, rootObject.EX_PRD_DET.MATNR));
                                EntityCollection EntityProductColl = _crmService.RetrieveMultiple(query);
                                if (EntityProductColl.Entities.Count > 0)
                                {
                                    string ProductId = EntityProductColl.Entities[0].Id.ToString();
                                    string ProductNumber = EntityProductColl.Entities[0].Contains("productnumber") ? EntityProductColl.Entities[0].GetAttributeValue<string>("productnumber") : "";
                                    EntityReference? MaterialGroup = EntityProductColl.Entities[0].Contains("hil_materialgroup") ? EntityProductColl.Entities[0].GetAttributeValue<EntityReference>("hil_materialgroup") : null;
                                    EntityReference? Division = EntityProductColl.Entities[0].Contains("hil_division") ? EntityProductColl.Entities[0].GetAttributeValue<EntityReference>("hil_division") : null;
                                    string Description = EntityProductColl.Entities[0].Contains("description") ? EntityProductColl.Entities[0].GetAttributeValue<string>("description") : "";
                                    if (Division == null || MaterialGroup == null)
                                    {
                                        return new ValidateSerialNumResponse
                                        {
                                            SerialNumber = SerialNumber,
                                            StatusCode = 204,
                                            Message = CommonMessage.Division_materialgroupMsg,
                                            //Message = "Division or Material group mapping is missing."
                                        };
                                    }
                                    string Brand = "";
                                    Entity _entBrand = _crmService.Retrieve("product", Division.Id, new ColumnSet("hil_brandidentifier"));
                                    if (_entBrand != null)
                                    {
                                        Brand = _entBrand.FormattedValues["hil_brandidentifier"].ToString();
                                    }
                                    result = new ValidateSerialNumResponse
                                    {
                                        ProductId = ProductId,
                                        ModelCode = ProductNumber,
                                        ModelName = Description,
                                        ProductCategory = Division.Name,
                                        ProductCategoryGuid = Division.Id,
                                        ProductSubCategory = MaterialGroup.Name,
                                        ProductSubCategoryGuid = MaterialGroup.Id,
                                        SerialNumber = SerialNumber,
                                        Brand = Brand,
                                        PTAStatus = true,
                                        StatusCode = 200,
                                        Message = "Success",
                                        TokenExpiresAt = expiryTime,
                                    };
                                }
                                else
                                {
                                    result = new ValidateSerialNumResponse
                                    {
                                        PTAStatus = false,
                                        SerialNumber = SerialNumber,
                                        StatusCode = 204,
                                        Message = CommonMessage.Modelnotexist,
                                        //Message = "Model does't exist in D365.",
                                        TokenExpiresAt = expiryTime

                                    };
                                }
                            }
                            else
                            {
                                result = new ValidateSerialNumResponse
                                {
                                    PTAStatus = false,
                                    SerialNumber = SerialNumber,
                                    StatusCode = 204,
                                    Message = CommonMessage.Serialnumberalreadyexist,
                                    // Message = "Provided Serial Number (IDU) already exists.",
                                    TokenExpiresAt = expiryTime
                                };
                            }
                        }
                        else
                        {
                            result = new ValidateSerialNumResponse
                            {
                                PTAStatus = false,
                                SerialNumber = SerialNumber,
                                StatusCode = 204,
                                Message = CommonMessage.SAPapiConfigMsg,
                                //Message = "SAP API Config not found.",
                                TokenExpiresAt = expiryTime
                            };
                        }

                        return result;
                    }

                    return new ValidateSerialNumResponse
                    {
                        PTAStatus = true,
                        SerialNumber = SerialNumber,
                        StatusCode = 204,
                        Message = CommonMessage.Serialnumberalreadyexist,
                        //Message = "Serial Number already exist.",
                        TokenExpiresAt = expiryTime
                    };
                }
                return new ValidateSerialNumResponse
                {
                    StatusCode = 503,
                    Message = CommonMessage.ServiceUnavailableMsg
                    // Message = "D365 service unavailable."
                };
            }
            catch (Exception ex)
            {
                return new ValidateSerialNumResponse
                {
                    StatusCode = 500,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message
                    // Message = "D365 internal server error : " + ex.Message.ToUpper()
                };
            }
        }
        public bool CheckIfExistingSerialNumber(ICrmService service, string Serial)
        {
            bool flag = true;
            QueryExpression queryExpression = new QueryExpression("msdyn_customerasset");
            queryExpression.ColumnSet = new ColumnSet(allColumns: false);
            queryExpression.Criteria = new FilterExpression(LogicalOperator.And);
            queryExpression.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, Serial);
            EntityCollection entityCollection = service.RetrieveMultiple(queryExpression);
            if (entityCollection.Entities.Count > 0)
            {
                return true;
            }

            return false;
        }

        public Entity CheckIfExistingSerialNumberWithDetails(ICrmService service, string Serial)
        {
            Entity result = null;
            QueryExpression queryExpression = new QueryExpression("msdyn_customerasset");
            queryExpression.ColumnSet = new ColumnSet("msdyn_name");
            queryExpression.Criteria = new FilterExpression(LogicalOperator.And);
            queryExpression.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, Serial);
            EntityCollection entityCollection = service.RetrieveMultiple(queryExpression);
            if (entityCollection.Entities.Count > 0)
            {
                result = entityCollection.Entities[0];
            }
            return result;
        }

        public async Task<RegisterProductResponse> RegisterProduct(RegisterProductModel objProductData, string SessionId)
        {
            var objResult = new RegisterProductResponse();
            DateTime InvoiceDate;
            string[] formats = { "yyyy-MM-dd" };
            byte[] NoteByte = null;
            try
            {
                if (_crmService != null)
                {
                    string expiryTime = CommonMethods.getExpiryTime(_crmService, SessionId).ToString();
                    if (objProductData.ProductId == Guid.Empty || CommonMethods.IsvalidGuid(_crmService, "product", objProductData.ProductId) == false)
                    {
                        return new RegisterProductResponse
                        {
                            StatusCode = 204,
                            Message = CommonMessage.InvalidProductGuidMsg,
                            //Message = "Invalid Product Guid.",
                            TokenExpiresAt = expiryTime
                        };
                    }
                    if (string.IsNullOrEmpty(objProductData.SerialNumber))
                    {
                        return new RegisterProductResponse
                        {
                            StatusCode = 204,
                            Message = CommonMessage.SerialnumberMsg,
                            //Message = "Serial Number is required.",
                            TokenExpiresAt = expiryTime
                        };
                    }
                    if (objProductData.CustomerGuid == Guid.Empty || CommonMethods.IsvalidGuid(_crmService, "contact", objProductData.CustomerGuid) == false)
                    {
                        return new RegisterProductResponse
                        {
                            StatusCode = 204,
                            Message = CommonMessage.InvalidCustomerGuid,
                            //Message = "Invalid Customer Guid.",
                            TokenExpiresAt = expiryTime
                        };
                    }
                    if (!DateTime.TryParseExact(objProductData.InvoiceDate, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out InvoiceDate))
                    {
                        return new RegisterProductResponse
                        {
                            StatusCode = 204,
                            Message = CommonMessage.InvalidInvoiceDateMsg,
                            // Message = "Invalid Invoice Date. Please try with this format(yyyy-MM-dd).",
                            TokenExpiresAt = expiryTime
                        };
                    }

                    if (!string.IsNullOrEmpty(objProductData.InvoiceBase64))
                    {
                        try
                        {
                            NoteByte = Convert.FromBase64String(objProductData.InvoiceBase64);
                        }
                        catch (Exception ex)
                        {
                            return new RegisterProductResponse
                            {
                                StatusCode = 204,
                                Message = CommonMessage.InvalidinvoiceBase64,
                                //Message = "Invalid invoiceBase64.",
                                TokenExpiresAt = expiryTime
                            };
                        }
                        if (objProductData.FileType < 0 || objProductData.FileType > 6)
                        {
                            return new RegisterProductResponse
                            {
                                StatusCode = 204,
                                Message = CommonMessage.InvalidFileType,
                                // Message = "Invalid File Type.",
                                TokenExpiresAt = expiryTime
                            };
                        }
                    }
                    QueryExpression qry = new QueryExpression("msdyn_customerasset");
                    qry.ColumnSet = new ColumnSet(false);
                    qry.Criteria = new FilterExpression(LogicalOperator.And);
                    qry.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, objProductData.SerialNumber));
                    EntityCollection Coll = _crmService.RetrieveMultiple(qry);
                    if (Coll.Entities.Count > 0)
                    {
                        return new RegisterProductResponse
                        {
                            StatusCode = 204,
                            Message = CommonMessage.Serialnumberalreadyexist,
                            //Message = "Provided Serial Number (IDU) already exists.",
                            TokenExpiresAt = expiryTime
                        };
                    }
                    Entity _CstA = new Entity("msdyn_customerasset");
                    _CstA["msdyn_name"] = objProductData.SerialNumber;
                    _CstA["hil_customer"] = new EntityReference("contact", objProductData.CustomerGuid);
                    _CstA["msdyn_product"] = new EntityReference("product", objProductData.ProductId);
                    _CstA["hil_invoicedate"] = InvoiceDate;
                    _CstA["hil_invoiceavailable"] = true;

                    Entity ent = _crmService.Retrieve("product", objProductData.ProductId, new ColumnSet("hil_division", "hil_materialgroup", "description", "name", "hil_brandidentifier"));
                    EntityReference? ProductCategoryId = ent.Attributes.Contains("hil_division") ? ent.GetAttributeValue<EntityReference>("hil_division") : null;
                    EntityReference? ProductSubcategoryId = ent.Attributes.Contains("hil_materialgroup") ? ent.GetAttributeValue<EntityReference>("hil_materialgroup") : null;
                    string Brand = ent.Attributes.Contains("hil_brandidentifier") ? ent.FormattedValues["hil_brandidentifier"] : "";
                    Entity _entBrand = _crmService.Retrieve("product", ProductCategoryId.Id, new ColumnSet("hil_brandidentifier"));
                    if (_entBrand != null)
                    {
                        Brand = _entBrand.FormattedValues["hil_brandidentifier"].ToString();
                    }
                    string Description = ent.Attributes.Contains("description") ? ent.GetAttributeValue<string>("description") : "";

                    if (ProductCategoryId != null && ProductSubcategoryId != null)
                    {
                        _CstA["hil_productcategory"] = ProductCategoryId;
                        _CstA["hil_productsubcategory"] = ProductSubcategoryId;

                        qry = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
                        qry.ColumnSet = new ColumnSet(false);
                        qry.Criteria = new FilterExpression(LogicalOperator.And);
                        qry.Criteria.AddCondition(new ConditionExpression("hil_productcategorydivision", ConditionOperator.Equal, ProductCategoryId.Id));
                        qry.Criteria.AddCondition(new ConditionExpression("hil_productsubcategorymg", ConditionOperator.Equal, ProductSubcategoryId.Id));
                        Coll = _crmService.RetrieveMultiple(qry);
                        if (Coll.Entities.Count > 0)
                        {
                            _CstA["hil_productsubcategorymapping"] = new EntityReference("hil_stagingdivisonmaterialgroupmapping", Coll.Entities[0].Id);
                        }
                    }
                    else
                    {
                        return new RegisterProductResponse
                        {
                            StatusCode = 503,
                            Message = CommonMessage.Division_materialgroupMsg
                            // Message = "Division and Material Group mapping is unavailable."
                        };
                    }
                    Entity Consumerloginsession = _crmService.Retrieve("hil_consumerloginsession", new Guid(SessionId), new ColumnSet("hil_origin"));
                    int source = Convert.ToInt32(Consumerloginsession.GetAttributeValue<string>("hil_origin"));

                    _CstA["hil_source"] = new OptionSetValue(source); // AMC Sale - SYNC APP
                    _CstA["statuscode"] = new OptionSetValue(910590000); // Pending for Approval
                    Guid CustomerAssetId = Guid.Empty;
                    try
                    {
                        CustomerAssetId = _crmService.Create(_CstA);
                    }
                    catch (Exception ex)
                    {
                        return new RegisterProductResponse
                        {
                            StatusCode = 204,
                            Message = ex.Message,
                            TokenExpiresAt = expiryTime
                        };
                    }

                    if (CustomerAssetId != Guid.Empty)
                    {
                        if (!string.IsNullOrEmpty(objProductData.InvoiceBase64))
                        {
                            CommonMethods.AttachNotes(_crmService, NoteByte, CustomerAssetId, objProductData.FileType, "msdyn_customerasset");
                        }
                        return new RegisterProductResponse
                        {
                            StatusCode = 200,
                            Message = CommonMessage.SuccessMsg,
                            Brand = Brand,
                            ProductDesc = Description,
                            CustomerAssetGuid = CustomerAssetId,
                            TokenExpiresAt = expiryTime
                        };
                    }
                    return objResult;
                }
                else
                {
                    return new RegisterProductResponse
                    {
                        StatusCode = 503,
                        Message = CommonMessage.ServiceUnavailableMsg,
                        //Message = "D365 service unavailable."
                    };
                }
            }
            catch (Exception ex)
            {
                return new RegisterProductResponse
                {
                    StatusCode = 500,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message
                    //Message = "D365 internal server error : " + ex.Message.ToUpper()
                };
            }
        }
        public async Task<ProductDeatilsList> GetRegisteredProducts(RegisterProductModel registeredProduct, string SessionId)
        {
            ProductDetails objRegisteredProducts;
            ProductDeatilsList lstRegisteredProducts = new ProductDeatilsList();
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                if (_crmService != null)
                {
                    string expiryTime = CommonMethods.getExpiryTime(_crmService, SessionId).ToString();

                    if (registeredProduct.CustomerGuid == Guid.Empty)
                    {
                        lstRegisteredProducts.StatusCode = 204;
                        lstRegisteredProducts.Message = CommonMessage.CustomerguidMsg; //"Customer Guid is required.";
                        lstRegisteredProducts.TokenExpiresAt = expiryTime;
                        return lstRegisteredProducts;
                    }

                    Query = new QueryExpression("msdyn_customerasset");
                    Query.ColumnSet = new ColumnSet("hil_warrantytilldate", "hil_warrantysubstatus", "hil_warrantystatus", "hil_modelname", "hil_product", "hil_retailerpincode", "hil_purchasedfrom", "hil_invoicevalue", "hil_invoiceno", "hil_invoicedate", "hil_invoiceavailable", "hil_batchnumber", "hil_pincode", "msdyn_customerassetid", "createdon", "msdyn_product", "msdyn_name", "hil_productsubcategorymapping", "hil_productsubcategory", "hil_productcategory", "hil_source");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, registeredProduct.CustomerGuid);
                    Query.AddOrder("createdon", OrderType.Descending);

                    entcoll = _crmService.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        lstRegisteredProducts.StatusCode = 204;
                        lstRegisteredProducts.Message = CommonMessage.NoRecordFound;
                        // lstRegisteredProducts.Message = "No Product found against the user.";
                    }
                    else
                    {
                        lstRegisteredProducts.ProductList = new List<ProductDetails>();
                        foreach (Entity ent in entcoll.Entities)
                        {
                            objRegisteredProducts = new ProductDetails();
                            objRegisteredProducts.CustomerGuid = registeredProduct.CustomerGuid;
                            objRegisteredProducts.ProductId = ent.GetAttributeValue<Guid>("msdyn_customerassetid");

                            if (ent.Attributes.Contains("hil_productcategory"))
                            {
                                objRegisteredProducts.ProductCategory = ent.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                                objRegisteredProducts.ProductCategoryId = ent.GetAttributeValue<EntityReference>("hil_productcategory").Id;
                            }
                            if (ent.Attributes.Contains("msdyn_product"))
                            {
                                objRegisteredProducts.ModelCode = ent.GetAttributeValue<EntityReference>("msdyn_product").Name;
                                objRegisteredProducts.ProductGuid = ent.GetAttributeValue<EntityReference>("msdyn_product").Id;

                                Entity EntProduct = _crmService.Retrieve("product", ent.GetAttributeValue<EntityReference>("msdyn_product").Id, new ColumnSet("description", "hil_brandidentifier"));
                                //objRegisteredProducts.ProductDesc = EntProduct.Contains("description") ? EntProduct.GetAttributeValue<string>("description") : "";
                                objRegisteredProducts.Brand = EntProduct.FormattedValues.Contains("hil_brandidentifier") ? EntProduct.FormattedValues["hil_brandidentifier"] : "";
                                Entity _entBrand = _crmService.Retrieve("product", objRegisteredProducts.ProductCategoryId, new ColumnSet("hil_brandidentifier"));
                                if (_entBrand != null)
                                {
                                    objRegisteredProducts.Brand = _entBrand.FormattedValues["hil_brandidentifier"].ToString();
                                }

                            }
                            //objRegisteredProducts.ProductWarranty = new List<ProductWarranty>();
                            objRegisteredProducts.ProductWarranty = GetWarrantyDetails(objRegisteredProducts.ProductId, _crmService);
                            if (ent.Attributes.Contains("hil_modelname"))
                            { objRegisteredProducts.ProductDesc = ent.GetAttributeValue<string>("hil_modelname"); }
                            if (ent.Attributes.Contains("hil_productsubcategory"))
                            {
                                objRegisteredProducts.ProductSubCategory = ent.GetAttributeValue<EntityReference>("hil_productsubcategory").Name;
                                objRegisteredProducts.ProductSubCategoryId = ent.GetAttributeValue<EntityReference>("hil_productsubcategory").Id;
                            }
                            if (ent.Attributes.Contains("msdyn_name"))
                            { objRegisteredProducts.SerialNumber = ent.GetAttributeValue<string>("msdyn_name"); }
                            //if (ent.Attributes.Contains("hil_batchnumber"))
                            //{ objRegisteredProducts.BatchNumber = ent.GetAttributeValue<string>("hil_batchnumber"); }
                            if (ent.Attributes.Contains("hil_invoiceavailable"))
                            { objRegisteredProducts.InvoiceAvailable = ent.GetAttributeValue<bool>("hil_invoiceavailable"); }
                            if (ent.Attributes.Contains("hil_invoiceno"))
                            { objRegisteredProducts.InvoiceNumber = ent.GetAttributeValue<string>("hil_invoiceno"); }
                            if (ent.Attributes.Contains("hil_invoicedate"))
                            { objRegisteredProducts.InvoiceDate = ent.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330).ToString("dd/MM/yyyy"); }
                            if (ent.Attributes.Contains("hil_invoicevalue"))
                            { objRegisteredProducts.InvoiceValue = decimal.Round(ent.GetAttributeValue<decimal>("hil_invoicevalue"),2); }

                            if (ent.Attributes.Contains("hil_purchasedfrom"))
                            { objRegisteredProducts.PurchasedFrom = ent.GetAttributeValue<string>("hil_purchasedfrom"); }
                            if (ent.Attributes.Contains("hil_retailerpincode"))
                            { objRegisteredProducts.PurchasedFromLocation = ent.GetAttributeValue<string>("hil_retailerpincode"); }

                            //if (ent.Attributes.Contains("hil_product"))
                            //{ objRegisteredProducts.InstalledLocationEnum = ent.GetAttributeValue<OptionSetValue>("hil_product").Value; }

                            //if (ent.Attributes.Contains("hil_product"))
                            //{ objRegisteredProducts.InstalledLocation = ent.FormattedValues["hil_product"].ToString(); }

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
                                objRegisteredProducts.WarrantyEndDate = ent.GetAttributeValue<DateTime>("hil_warrantytilldate").AddMinutes(330).ToString("dd/MM/yyyy");
                            }
                            objRegisteredProducts.SourceOfRegistration = ent.Contains("hil_source") ? ent.FormattedValues["hil_source"] : "";
                            lstRegisteredProducts.ProductList.Add(objRegisteredProducts);
                        }
                    }
                    lstRegisteredProducts.TokenExpiresAt = expiryTime;
                    lstRegisteredProducts.StatusCode = 200;
                    lstRegisteredProducts.Message = CommonMessage.SuccessMsg;       //"Success";
                }
                else
                {
                    lstRegisteredProducts.StatusCode = 503;
                    lstRegisteredProducts.Message = CommonMessage.ServiceUnavailableMsg;        //"D365 service unavailable.";
                    return lstRegisteredProducts;
                }
            }
            catch (Exception ex)
            {
                return new ProductDeatilsList
                {
                    StatusCode = 500,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message
                    //Message = "D365 internal server error : " + ex.Message.ToUpper()
                };
            }
            return lstRegisteredProducts;
        }

        private List<ProductWarranty> GetWarrantyDetails(Guid ProductGuid, ICrmService service)
        {
            List<ProductWarranty> lstWarrantyLineInfo = new List<ProductWarranty>();
            QueryExpression Query = new QueryExpression("hil_unitwarranty");
            Query.ColumnSet = new ColumnSet("hil_warrantystartdate", "hil_warrantyenddate", "hil_warrantytemplate");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, ProductGuid);
            Query.Criteria.AddCondition("hil_warrantyenddate", ConditionOperator.GreaterEqual, DateTime.Now);
            EntityCollection entColl = service.RetrieveMultiple(Query);

            if (entColl.Entities.Count > 0)
            {
                foreach (Entity entwarranty in entColl.Entities)
                {
                    ProductWarranty objWarrantyLineInfo = new ProductWarranty();

                    DateTime hil_warrantystartdate = DateTime.Now;
                    if (entwarranty.Contains("hil_warrantystartdate"))
                        hil_warrantystartdate = entwarranty.GetAttributeValue<DateTime>("hil_warrantystartdate");

                    DateTime hil_warrantyenddate = DateTime.Now;
                    if (entwarranty.Contains("hil_warrantyenddate"))
                        hil_warrantyenddate = entwarranty.GetAttributeValue<DateTime>("hil_warrantyenddate");

                    objWarrantyLineInfo.WarrantyStartDate = hil_warrantystartdate.ToString();
                    objWarrantyLineInfo.WarrantyEndDate = hil_warrantyenddate.ToString();

                    Guid Guidwarrantytemplate = Guid.Empty;
                    if (entwarranty.Contains("hil_warrantytemplate"))
                        Guidwarrantytemplate = entwarranty.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id;

                    if (Guidwarrantytemplate != Guid.Empty)
                    {
                        Entity entwarrantytemplate = service.Retrieve("hil_warrantytemplate", Guidwarrantytemplate, new ColumnSet("hil_warrantyperiod", "hil_type", "hil_description"));

                        if (entwarrantytemplate.Attributes.Count > 0)
                        {
                            string WarrantyPeriod = string.Empty;
                            if (entwarrantytemplate.Contains("hil_warrantyperiod"))
                                WarrantyPeriod = entwarrantytemplate.GetAttributeValue<Int32>("hil_warrantyperiod").ToString();

                            string WarrantyDescription = string.Empty;
                            if (entwarrantytemplate.Contains("hil_description"))
                                WarrantyDescription = entwarrantytemplate.GetAttributeValue<string>("hil_description");

                            string WarrantyType = string.Empty;
                            if (entwarrantytemplate.Contains("hil_type"))
                            {
                                WarrantyType = entwarrantytemplate.FormattedValues["hil_type"].ToString();
                            }
                            objWarrantyLineInfo.WarrantySpecifications = WarrantyDescription;
                            //objWarrantyLineInfo.WarrantyCoverage = WarrantyDescription;
                            //objWarrantyLineInfo.WarrantyPeriod = WarrantyPeriod;
                            objWarrantyLineInfo.WarrantyType = WarrantyType;
                        }
                    }
                    lstWarrantyLineInfo.Add(objWarrantyLineInfo);
                }
            }
            return lstWarrantyLineInfo;
        }
    }
}