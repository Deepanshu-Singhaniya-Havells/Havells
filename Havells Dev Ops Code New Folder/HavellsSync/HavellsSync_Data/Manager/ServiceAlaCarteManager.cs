using HavellsSync_Data.IServiceAlaCarte;
using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.ICommon;
using HavellsSync_ModelData.ServiceAlaCarte;
using Microsoft.Extensions.Configuration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Net;
using System.Net.Http.Headers;
using System.Text;

namespace HavellsSync_Data.Manager
{
    public class ServiceAlaCarteManager : IServiceAlaCarteManager
    {
        private IConfiguration configuration;
        private ICrmService _CrmService;
        private readonly string _PayNowKey;
        private readonly string _Surl;
        private readonly string _Furl;
        private readonly int _Waittime;
        private readonly string _EngMobileNo;
        private readonly ICustomLog _logger;
        private static EntityReference _SellingSource = new EntityReference("hil_integrationsource", new Guid("834acd00-cf94-ee11-be36-6045bd72672b")); // Havells ONE APP

        public ServiceAlaCarteManager(ICrmService crmService, IConfiguration configuration, ICustomLog logger)
        {
            Check.Argument.IsNotNull(nameof(crmService), crmService);
            _CrmService = crmService;
            this.configuration = configuration;
            _PayNowKey = configuration.GetSection(ConfigKeys.CRMSettings + ":" + ConfigKeys.PayNowKey).Value;
            _Surl = configuration.GetSection(ConfigKeys.CRMSettings + ":" + ConfigKeys.Surl).Value;
            _Furl = configuration.GetSection(ConfigKeys.CRMSettings + ":" + ConfigKeys.Furl).Value;
            _Waittime = Convert.ToInt32(configuration.GetSection(ConfigKeys.CRMSettings + ":" + ConfigKeys.Waittime).Value);
            _EngMobileNo = configuration.GetSection(ConfigKeys.CRMSettings + ":" + ConfigKeys.EngMobileNo).Value;
            _logger = logger;
        }
        public async Task<(List<ProductCatagories>, RequestStatus)> GetServiceProductCategory()
        {
            List<ProductCatagories> objCategorylist = new List<ProductCatagories>();
            try
            {
                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                <entity name='product'>
                                <attribute name='name' />
                                <attribute name='productid' />
                                <order attribute='productnumber' descending='false' />
                                <filter type='and'>
                                    <condition attribute='hil_hierarchylevel' operator='eq' value='2' />
                                </filter>
                                <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='ad'>
                                    <attribute name='hil_productimage' />
                                     <attribute name='hil_name' />
                                    <filter type='and'>
                                    <condition attribute='hil_enableforalacarte' operator='eq' value='1' />
                                    </filter>
                                </link-entity>
                                </entity>
                                </fetch>";

                EntityCollection entColl = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));

                if (entColl.Entities.Count > 0)
                {
                    foreach (Entity entity in entColl.Entities)
                    {
                        List<EntityInfo> lstEntityInfo = GetProductSubCategory(entity.Id);
                        ProductCatagories objitem = new ProductCatagories
                        {
                            ProductCategory = entity.Contains("ad.hil_name") ? entity.GetAttributeValue<AliasedValue>("ad.hil_name").Value.ToString() : "",
                            ProductCategoryId = entity.Contains("productid") ? entity.GetAttributeValue<Guid>("productid").ToString() : "",
                            ProductSubCategory = lstEntityInfo.Any() ? lstEntityInfo.Select(m => m.Name).FirstOrDefault() : "",
                            ProductSubCategoryId = lstEntityInfo.Any() ? lstEntityInfo.Select(m => m.Id).FirstOrDefault().ToString() : "",
                            CategoryImageUrl = entity.Contains("ad.hil_productimage") ? entity.GetAttributeValue<AliasedValue>("ad.hil_productimage").Value.ToString() : "",
                        };
                        objCategorylist.Add(objitem);
                    }
                }
                return (objCategorylist, new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.OK
                });
            }
            catch (Exception ex)
            {
                return (objCategorylist, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        private List<EntityInfo> GetProductSubCategory(Guid Id)
        {
            List<EntityInfo> lstEntityInfo = new List<EntityInfo>();
            string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='product'>
                                    <attribute name='name'/>
                                    <attribute name='productnumber'/>
                                    <attribute name='description'/>
                                    <attribute name='productid'/>
                                    <order attribute='name' descending='false'/>
                                    <filter type='and'>
                                        <condition attribute='hil_hierarchylevel' operator='eq' value='3'/>
                                        <condition attribute='hil_division' operator='eq' value='{Id}'/>
                                        <condition attribute='statecode' operator='eq' value='0'/>
                                    </filter>
                                    <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='ad'>
                                    <filter type='and'>
                                        <condition attribute='hil_enableforalacarte' operator='eq' value='1' />
                                    </filter>
                                    </link-entity>
                                    </entity>
                                    </fetch>";

            EntityCollection entColl = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));
            if (entColl.Entities.Count > 0)
            {
                foreach (Entity ent in entColl.Entities)
                {
                    EntityInfo entityInfo = new EntityInfo()
                    {
                        Id = ent.Id,
                        Name = ent.Contains("name") ? ent.GetAttributeValue<string>("name") : "",
                    };
                    lstEntityInfo.Add(entityInfo);
                }
            }
            return lstEntityInfo;
        }
        public async Task<(ServiceAlaCartePlanInfo, RequestStatus)> GetServiceAlaCarteList(string ProductSubCategoryId)
        {
            ServiceAlaCartePlanInfo serviceAlaCartePlanInfo = new ServiceAlaCartePlanInfo();
            serviceAlaCartePlanInfo.ServiceAlaCartePlanList = new List<ServiceAlaCartePlanList>();
            try
            {
                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                    <entity name='product'>
                                    <attribute name='productid' />
                                    <attribute name='hil_materialgroup' />
                                    <attribute name='hil_amount' />
                                    <order attribute='productnumber' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='hil_hierarchylevel' operator='eq' value='910590002' />
                                      <condition attribute='hil_materialgroup' operator='eq' value='{ProductSubCategoryId}' />
                                      <condition attribute='statuscode' operator='eq' value='1'/>
                                    </filter>
                                    <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='ah'>
                                    <attribute name='hil_productbannerimageurl' />
                                    <attribute name='hil_planperiod' />
                                    <attribute name='hil_notcovered' />
                                    <attribute name='hil_coverage' />
                                    <attribute name='hil_name' />
                                    <filter type='and'>
                                       <condition attribute='hil_enableforalacarte' operator='eq' value='1' />
                                    </filter>
                                    </link-entity>
                                    <link-entity name='productpricelevel' from='productid' to='productid' link-type='inner' alias='aa'>
                                        <attribute name='amount' />
                                        <attribute name='hil_discount' />
                                    </link-entity>
                                    </entity>
                                    </fetch>";

                EntityCollection Info1 = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));

                if (Info1.Entities.Count > 0)
                {
                    List<ServiceAlaCartePlanList> lstServiceAlaCartePlanList = new List<ServiceAlaCartePlanList>();
                    foreach (Entity entity in Info1.Entities)
                    {
                        ServiceAlaCartePlanList objitem = new ServiceAlaCartePlanList
                        {
                            SubCategoryId = entity.Contains("hil_materialgroup") ? entity.GetAttributeValue<EntityReference>("hil_materialgroup").Id : Guid.Empty,
                            SubCategoryName = entity.Contains("hil_materialgroup") ? entity.GetAttributeValue<EntityReference>("hil_materialgroup").Name.ToString() : null,
                            ServiceId = entity.Contains("productid") ? entity.GetAttributeValue<Guid>("productid").ToString() : null,
                            ServiceName = entity.Contains("ah.hil_name") ? entity.GetAttributeValue<AliasedValue>("ah.hil_name").Value.ToString() : null,
                            ServiceDuration = entity.Contains("ah.hil_planperiod") ? entity.GetAttributeValue<AliasedValue>("ah.hil_planperiod").Value.ToString() : null,
                            MRP = entity.Contains("aa.amount") ? ((Money)entity.GetAttributeValue<AliasedValue>("aa.amount").Value).Value.ToString("F2") : null,
                            DiscountPercent = entity.Contains("aa.hil_discount") ? ((Money)entity.GetAttributeValue<AliasedValue>("aa.hil_discount").Value).Value.ToString("F2") : null,
                            Included = entity.Contains("ah.hil_coverage") ? entity.GetAttributeValue<AliasedValue>("ah.hil_coverage").Value.ToString() : null,
                            Excluded = entity.Contains("ah.hil_notcovered") ? entity.GetAttributeValue<AliasedValue>("ah.hil_notcovered").Value.ToString() : null,
                        };
                        lstServiceAlaCartePlanList.Add(objitem);
                    }
                    serviceAlaCartePlanInfo.ServiceAlaCartePlanList = lstServiceAlaCartePlanList;
                    string fetchXmlBanner = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='hil_productcatalog'>
                                            <attribute name='hil_productbannerimageurl' />
                                            <order attribute='hil_productbannerimageurl' descending='false' />
                                            <filter type='and'>
                                            <condition attribute='hil_productcode' operator='eq' value='{ProductSubCategoryId}' />
                                            </filter>
                                            </entity>
                                            </fetch>";
                    EntityCollection Info = _CrmService.RetrieveMultiple(new FetchExpression(fetchXmlBanner));
                    if (Info.Entities.Count > 0)
                    {
                        serviceAlaCartePlanInfo.ServiceBannerUrl = Info.Entities[0].Contains("hil_productbannerimageurl") ? Info.Entities[0].GetAttributeValue<string>("hil_productbannerimageurl") : "";
                    }
                }
                return (serviceAlaCartePlanInfo, new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.OK
                });
            }
            catch (Exception ex)
            {
                return (serviceAlaCartePlanInfo, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        public async Task<(ServiceOrder, RequestStatus)> PaymentRetry(PaymentRetryParam paymentRetryReq, string MobileNumber)
        {
            ServiceOrder objServiceOrder = new ServiceOrder();
            try
            {
                if (paymentRetryReq.PaymentType == "1")
                {
                    string fetchXmldetails = $@"<fetch version='1.0' output-format='xml-platform'  top='1'  mapping='logical' distinct='false'>
                                  <entity name='hil_paymentreceipt'>
                                  <attribute name='hil_paymentreceiptid'/>
                                  <attribute name='hil_transactionid'/>       
                                  <attribute name='hil_hashname'/>  
                                  <attribute name='hil_tokenexpireson'/> 
                                  <attribute name='hil_salt'/>  
                                  <attribute name='hil_key'/>  
                                  <attribute name='hil_memorandumcode'/>  
                                  <attribute name='hil_email'/>  
                                  <attribute name='hil_mobilenumber'/>  
                                  <attribute name='hil_amount'/>
                                  <order attribute='createdon' descending='true'/>
                                  <filter type='and'>
                                      <condition attribute='hil_orderid' operator='eq' uitype='salesorder' value='{paymentRetryReq.OrderId}'/>
                                      <condition attribute='hil_mobilenumber' operator='eq'  value='{MobileNumber}'/>
                                  </filter>
                                  </entity>
                                  </fetch>";
                    EntityCollection entColl = _CrmService.RetrieveMultiple(new FetchExpression(fetchXmldetails));
                    if (entColl.Entities.Count > 0)
                    {
                        var Expiredt = entColl.Entities[0].GetAttributeValue<DateTime>("hil_tokenexpireson");

                        if (Expiredt >= DateTime.UtcNow)
                        {
                            return (objServiceOrder, new RequestStatus()
                            {
                                StatusCode = (int)HttpStatusCode.BadRequest,
                                Message = $"Your transaction is pending! Please wait and Try again after  {_Waittime} minutes.",
                            });
                        }
                        else
                        {
                            var result = Paynow("", new Guid(paymentRetryReq.OrderId), "", 0, "", "", "");

                            if (result.Item2.StatusCode == 200)
                            {
                                Guid receiptId = entColl.Entities[0].GetAttributeValue<Guid>("hil_paymentreceiptid");
                                if (!string.IsNullOrWhiteSpace(receiptId.ToString()) && receiptId != Guid.Empty)
                                {
                                    Entity paymentreceiptEnt = new Entity("hil_paymentreceipt", receiptId);
                                    paymentreceiptEnt["hil_paymentstatus"] = new OptionSetValue(2);
                                    _CrmService.Update(paymentreceiptEnt);
                                }
                                ServiceOrderDetails serviceOrderDetails = getServiceOrderDetails(new Guid(paymentRetryReq.OrderId), ref objServiceOrder);
                                objServiceOrder.PaymentInfo = result.Item1;
                            }
                        }
                        return (objServiceOrder, new RequestStatus
                        {
                            StatusCode = (int)HttpStatusCode.OK
                        });
                    }
                    else
                    {
                        return (objServiceOrder, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = "Invalid OrderId",
                        });
                    }
                }
                else
                {
                    return (objServiceOrder, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = "Invalid Payment Type",
                    });
                }
            }
            catch (Exception ex)
            {
                return (objServiceOrder, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        public async Task<(List<BookedService>, RequestStatus)> GetMostBookedServices()
        {
            List<BookedService> objCategorylist = new List<BookedService>();
            try
            {
                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='hil_productcatalog'>
                                      <attribute name='hil_productcatalogid' />
                                     <attribute name='hil_productcode' />
                                      <attribute name='hil_name' />
                                      <attribute name='hil_productbannerimageurl' />
                                      <order attribute='hil_omnichanneldisplayindex' descending='false' />
                                      <filter type='and'>
                                        <condition attribute='hil_enableforalacarte' operator='eq' value='1' />
                                        <condition attribute='hil_omnichanneldisplayindex' operator='gt' value='0' />
                                      </filter>
                                      <link-entity name='product' from='productid' to='hil_productcode' link-type='inner' alias='pc'>
                                        <attribute name='name' />
                                        <attribute name='description' />
                                        <attribute name='hil_division' />
                                          <filter type='and'>
                                              <condition attribute='hil_hierarchylevel' operator='eq' value='910590002' />
                                          </filter>
                                      <link-entity name='productpricelevel' from='productid' to='productid' link-type='inner' alias='aa'>
                                        <attribute name='amount' />
                                        <attribute name='hil_discount' />
                                      </link-entity>
                                      </link-entity>
                                      </entity>
                                      </fetch>";

                EntityCollection entColl = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));

                if (entColl.Entities.Count > 0)
                {
                    foreach (Entity entity in entColl.Entities)
                    {
                        BookedService objitem = new BookedService();
                        Money MRP = entity.Contains("aa.amount") ? (Money)entity.GetAttributeValue<AliasedValue>("aa.amount").Value : new Money(0);
                        if (MRP.Value == 0)
                        {
                            continue;
                        }
                        if (entity.Contains("pc.hil_division"))
                        {
                            EntityReference CategoryEnt = (EntityReference)entity.GetAttributeValue<AliasedValue>("pc.hil_division").Value;

                            fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='hil_productcatalog'>
                                            <attribute name='hil_productcatalogid' />
                                            <attribute name='hil_name' />
                                            <attribute name='createdon' />
                                            <order attribute='createdon' descending='true' />
                                            <filter type='and'>
                                              <condition attribute='hil_productcode' operator='eq' value='{CategoryEnt.Id}' />
                                            </filter>
                                            <link-entity name='product' from='productid' to='hil_productcode' link-type='inner' alias='ae'>
                                              <attribute name='hil_division' />
                                              <filter type='and'>
                                                <condition attribute='hil_hierarchylevel' operator='eq' value='2' />
                                              </filter>
                                            </link-entity>
                                          </entity>
                                        </fetch>";
                            EntityCollection entCollDivision = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));

                            if (entCollDivision.Entities.Count > 0)
                            {
                                objitem.ProductCategory = entCollDivision.Entities[0].Contains("hil_name") ? entCollDivision.Entities[0].GetAttributeValue<string>("hil_name") : null;
                            }
                            objitem.ProductCategoryId = CategoryEnt.Id.ToString();
                            List<EntityInfo> lstEntityInfo = GetProductSubCategory(CategoryEnt.Id);
                            objitem.ProductSubCategory = lstEntityInfo.Any() ? lstEntityInfo.Select(m => m.Name).FirstOrDefault() : "";
                            objitem.ProductSubCategoryId = lstEntityInfo.Any() ? lstEntityInfo.Select(m => m.Id).FirstOrDefault().ToString() : "";
                            objitem.ServiceId = entity.Contains("hil_productcode") ? entity.GetAttributeValue<EntityReference>("hil_productcode").Id.ToString() : null;
                            objitem.ServiceName = entity.Contains("hil_name") ? entity.GetAttributeValue<string>("hil_name") : null;
                            objitem.ServicePrice = entity.Contains("aa.amount") ? decimal.Round(MRP.Value, 2).ToString() : null;
                            objitem.ServiceDisplayIndex = entity.GetAttributeValue<Int32>("hil_omnichanneldisplayindex");
                            objitem.CategoryImageUrl = getImgUrl(CategoryEnt.Id);
                            objCategorylist.Add(objitem);
                        }
                    }
                }
                RequestStatus requestStatus = new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.OK,
                    Message = "Success",
                };
                return (objCategorylist, requestStatus);
            }
            catch (Exception ex)
            {
                return (objCategorylist, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        public async Task<(ServiceOrder, RequestStatus)> CreateServiceOrder(CreateOrder objOrder)
        {
            ServiceOrder objServiceOrder = new ServiceOrder();
            PaymentDetails objpayment = new PaymentDetails();
            objServiceOrder.BookedServiceList = new List<BookServies>();
            try
            {
                if (objOrder != null)
                {
                    string Cust_Name = string.Empty;
                    string Cust_email = string.Empty;
                    string Cust_mobile = string.Empty;
                    if (!await ValidateALaCartePriceAndDiscount(objOrder))
                    {
                        return (objServiceOrder, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = CommonMessage.PriceValidationMessage,
                        });
                    }
                    Entity AddressEnt = _CrmService.Retrieve("hil_address", new Guid(objOrder.AddressId), new ColumnSet("hil_addressid"));
                    if (AddressEnt == null)
                        return (objServiceOrder, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = CommonMessage.AddressNotExitMsg,
                        });
                    foreach (var item in objOrder.ServiceList)
                    {
                        if (!Validate.IsvalidGuid(item.ServiceId))
                            return (objServiceOrder, new RequestStatus()
                            {
                                StatusCode = (int)HttpStatusCode.BadRequest,
                                Message = CommonMessage.ServiceNotExits,
                            });
                        if (string.IsNullOrEmpty(item.ServiceName))
                            return (objServiceOrder, new RequestStatus()
                            {
                                StatusCode = (int)HttpStatusCode.BadRequest,
                                Message = CommonMessage.ServiceNameRequird,
                            });
                        if (!Validate.IsNumericGreaterThanZero(item.Quantity))
                            return (objServiceOrder, new RequestStatus()
                            {
                                StatusCode = (int)HttpStatusCode.BadRequest,
                                Message = CommonMessage.QuantityValid,
                            });
                        if (!Validate.IsNumericGreaterThanZero(item.MRP))
                            return (objServiceOrder, new RequestStatus()
                            {
                                StatusCode = (int)HttpStatusCode.BadRequest,
                                Message = CommonMessage.MRPValid
                            });
                        if (!Validate.IsNumeric(item.DiscountValue))
                            return (objServiceOrder, new RequestStatus()
                            {
                                StatusCode = (int)HttpStatusCode.BadRequest,
                                Message = CommonMessage.InvalidDiscount
                            });
                        if (!Validate.IsNumeric(item.DiscountPercent))
                            return (objServiceOrder, new RequestStatus()
                            {
                                StatusCode = (int)HttpStatusCode.BadRequest,
                                Message = CommonMessage.InvalidDiscount
                            });

                    }

                    Entity ContactEnt = _CrmService.Retrieve("contact", new Guid(objOrder.CustomerId), new ColumnSet("fullname", "emailaddress1", "mobilephone", "accountid"));
                    if (ContactEnt == null)
                    {
                        return (objServiceOrder, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = CommonMessage.CustomerNotExitMsg,
                        });
                    }
                    else
                    {
                        Cust_Name = ContactEnt.Contains("fullname") ? ContactEnt.GetAttributeValue<string>("fullname") : "";
                        Cust_email = ContactEnt.Contains("emailaddress1") ? ContactEnt.GetAttributeValue<string>("emailaddress1") : "";
                        Cust_mobile = ContactEnt.Contains("mobilephone") ? ContactEnt.GetAttributeValue<string>("mobilephone") : "";
                    }

                    Entity salesorderEnt = new Entity("salesorder");
                    salesorderEnt["msdyn_account"] = new EntityReference("account", new Guid("d166ba69-65da-ec11-a7b5-6045bdad2a19"));
                    salesorderEnt["customerid"] = new EntityReference("contact", new Guid(objOrder.CustomerId));
                    //salesorderEnt["totalamount_base"] = decimal.Round(Convert.ToDecimal(objOrder.OrderValue));
                    //salesorderEnt["discountamount"] = new Money(decimal.Round(Convert.ToDecimal(objOrder.DiscountAmount), 2)); // right(but value not save in entity) 
                    salesorderEnt["hil_receiptamount"] = new Money(decimal.Round(Convert.ToDecimal(objOrder.ReceiptAmount), 2)); // option(save value) 
                    salesorderEnt["totalamount"] = decimal.Round(Convert.ToDecimal(objOrder.ReceiptAmount), 2); //totallineitemamount,hil_receiptamount,totalamount,
                    salesorderEnt["hil_serviceaddress"] = new EntityReference("hil_address", new Guid(objOrder.AddressId));
                    salesorderEnt["requestdeliveryby"] = DateTime.Parse(objOrder.PreferredDate);
                    salesorderEnt["hil_preferreddaytime"] = new OptionSetValue(Convert.ToInt32(objOrder.PreferredDateTime));
                    salesorderEnt["msdyn_ordertype"] = new OptionSetValue(690970002);
                    salesorderEnt["hil_sellingsource"] = _SellingSource;
                    salesorderEnt["hil_modeofpayment"] = new OptionSetValue(Convert.ToInt32(objOrder.PaymentType));
                    salesorderEnt["hil_ordertype"] = new EntityReference("hil_ordertype", new Guid("019f761c-1669-ef11-a670-000d3a3e636d"));
                    salesorderEnt["msdyn_psastatusreason"] = new OptionSetValue(192350000);
                    salesorderEnt["transactioncurrencyid"] = new EntityReference("pricelevel", new Guid("68a6a9ca-6beb-e811-a96c-000d3af05828"));
                    salesorderEnt["pricelevelid"] = new EntityReference("pricelevel", new Guid("5688c4d2-bf70-ef11-a670-6045bdced00e"));
                    Guid orderID = _CrmService.Create(salesorderEnt);
                    if (orderID != Guid.Empty)
                    {
                        foreach (Services OrderLine in objOrder.ServiceList)
                        {
                            Entity entSalesOrderDetails = new Entity("salesorderdetail");
                            EntityReference ProductEnt = new EntityReference("product", new Guid(OrderLine.ServiceId));
                            entSalesOrderDetails["productid"] = ProductEnt;
                            entSalesOrderDetails["hil_product"] = ProductEnt;
                            entSalesOrderDetails["salesorderid"] = new EntityReference("salesorder", orderID);
                            entSalesOrderDetails["quantity"] = Convert.ToDecimal(OrderLine.Quantity);
                            //entSalesOrderDetails["baseamount"] = new Money(decimal.Round(Convert.ToDecimal(OrderLine.MRP), 2));  // right(but value not save in entity) 
                            //entSalesOrderDetails["priceperunit"] = new Money(decimal.Round(Convert.ToDecimal(OrderLine.MRP), 2)); // option(save value) 
                            entSalesOrderDetails["msdyn_billingmethod"] = new OptionSetValue(192350001); // Time and Material
                            entSalesOrderDetails["manualdiscountamount"] = objOrder.PaymentType == "1" ? new Money(decimal.Round(Convert.ToDecimal(OrderLine.DiscountValue), 2)) : new Money(0);
                            entSalesOrderDetails["msdyn_billingstatus"] = new OptionSetValue(objOrder.PaymentType == "1" ? 192350001 : 192350000);//Customer Invoice Created, Unbilled Sales Created
                                                                                                                                                  //entSalesOrderDetails["hil_discountpercent"] = OrderLine.DiscountPercent;
                            entSalesOrderDetails["uomid"] = new EntityReference("uom", new Guid("0359D51B-D7CF-43B1-87F6-FC13A2C1DEC8"));
                            _CrmService.Create(entSalesOrderDetails);
                        }
                    }
                    getServiceOrderDetails(orderID, ref objServiceOrder);
                    if (orderID != Guid.Empty)
                    {
                        if (objOrder.PaymentType == "1")
                        {
                            try
                            {
                                var result = Paynow(objOrder.CustomerId, orderID, objOrder.AddressId, Convert.ToDecimal(objOrder.ReceiptAmount), Cust_Name, Cust_mobile, Cust_email);
                                if (result.Item2.StatusCode == (int)HttpStatusCode.OK)
                                {
                                    objServiceOrder.PaymentInfo = result.Item1;
                                }
                                else
                                {
                                    return (new ServiceOrder(), new RequestStatus() { StatusCode = result.Item2.StatusCode, Message = result.Item2.Message });
                                }
                            }
                            catch (Exception ex)
                            {
                                return (new ServiceOrder(), new RequestStatus() { StatusCode = (int)HttpStatusCode.BadRequest, Message = ex.Message });
                            }
                        }
                    }
                    else
                    {
                        return (objServiceOrder, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = CommonMessage.OrderNotCreated,
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                return (objServiceOrder, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
            return (objServiceOrder, new RequestStatus()
            {
                StatusCode = (int)HttpStatusCode.OK,
                Message = CommonMessage.SuccessMsg,
            });
        }
        private ServiceOrderDetails getServiceOrderDetails(Guid orderID, ref ServiceOrder objServiceOrder)
        {
            ServiceOrderDetails serviceOrderDetails = new ServiceOrderDetails();
            Entity ServiceEnt = _CrmService.Retrieve("salesorder", orderID, new ColumnSet("hil_serviceaddress", "requestdeliveryby", "hil_preferreddaytime", "customerid", "hil_productdivision"));

            objServiceOrder.OrderId = orderID.ToString();
            Entity AddressCol = _CrmService.Retrieve("hil_address", ServiceEnt.GetAttributeValue<EntityReference>("hil_serviceaddress").Id, new ColumnSet("hil_addresstype"));
            objServiceOrder.AddressType = AddressCol.GetAttributeValue<OptionSetValue>("hil_addresstype").Value.ToString();
            objServiceOrder.PreferredDate = ServiceEnt.GetAttributeValue<DateTime>("requestdeliveryby").AddMinutes(330).ToString("dd/MM/yyyy");
            objServiceOrder.PreferredDateTime = ServiceEnt.GetAttributeValue<OptionSetValue>("hil_preferreddaytime").Value.ToString();
            if (ServiceEnt.Contains("hil_productdivision"))
            {
                EntityReference EntRefDivision = ServiceEnt.GetAttributeValue<EntityReference>("hil_productdivision");
                objServiceOrder.ProductCategory = EntRefDivision.Name;
                objServiceOrder.ProductCategoryId = EntRefDivision.Id.ToString();
                objServiceOrder.CategoryImageUrl = getImgUrl(EntRefDivision.Id);
            }
            QueryExpression queryExpression = new QueryExpression("salesorderdetail");
            queryExpression.ColumnSet = new ColumnSet("productid", "priceperunit", "quantity", "baseamount");
            queryExpression.Criteria.AddCondition("salesorderid", ConditionOperator.Equal, orderID);
            EntityCollection servicelineEntCol = _CrmService.RetrieveMultiple(queryExpression);
            List<BookServies> lstbookservice = new List<BookServies>();

            foreach (var item in servicelineEntCol.Entities)
            {
                BookServies objbookservice = new BookServies();
                Guid ProductCode = item.GetAttributeValue<EntityReference>("productid").Id;
                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' top='1' distinct='true'>
                                         <entity name='product'>
                                         <attribute name='name'/>
                                         <attribute name='description' />
                                         <attribute name='productid'/>
                                         <order attribute='productnumber' descending='false'/>
                                         <filter type='and'>
                                              <condition attribute='hil_hierarchylevel' operator='eq' value='910590002'/>
                                              <condition attribute='productid' operator='eq' value='{ProductCode}' />
                                         </filter>
                                         <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='ae'>
                                            <attribute name='hil_name' />
                                            <attribute name='hil_plantclink' />
                                             <filter type='and'>
                                                 <condition attribute='statuscode' operator='eq' value='1'/>
                                                 <condition attribute='hil_enableforalacarte' operator='eq' value='1'/>
                                             </filter>
                                         </link-entity>
                                         </entity>
                                         </fetch>";
                EntityCollection entitycol = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));

                if (entitycol.Entities.Count > 0)
                {
                    objbookservice.ServiceId = entitycol.Entities[0].Id.ToString();
                    objbookservice.ServiceName = entitycol.Entities[0].Contains("ae.hil_name") ? entitycol.Entities[0].GetAttributeValue<AliasedValue>("ae.hil_name").Value.ToString() : "";
                }
                if (item.Contains("priceperunit"))
                {
                    objbookservice.MRP = Convert.ToDouble(item.GetAttributeValue<Money>("priceperunit").Value.ToString());
                }
                //if (item.Contains("baseamount"))
                //{
                //    objbookservice.MRP = Convert.ToDouble(item.GetAttributeValue<Money>("baseamount").Value.ToString());
                //}
                objbookservice.Quantity = item.Contains("quantity") ? Convert.ToInt32(item.GetAttributeValue<decimal>("quantity")) : 0;

                lstbookservice.Add(objbookservice);
            }
            objServiceOrder.BookedServiceList = lstbookservice;

            serviceOrderDetails.serviceOrder = objServiceOrder;
            serviceOrderDetails.CustomerName = ServiceEnt.GetAttributeValue<EntityReference>("customerid").Name;
            return serviceOrderDetails;
        }
        private (PaymentDetails, RequestStatus) Paynow(string CustomerID, Guid orderid, string AddressID, decimal Amount, string CustomerName, string MobileNumber, string Email)
        {
            if (orderid != Guid.Empty && string.IsNullOrEmpty(CustomerID) && string.IsNullOrEmpty(MobileNumber))
            {
                string xmlquery = $@"<fetch>
                               <entity name='salesorder'>
                               <attribute name='salesorderid' />
                               <attribute name='customerid' />
                               <attribute name='hil_receiptamount' />
                               <attribute name='hil_serviceaddress' />
                               <attribute name='totalamount' />
                               <filter>
                               <condition attribute='salesorderid' operator='eq' value='{orderid}' />
                               </filter>
                               <link-entity name='contact' from='contactid' to='customerid' link-type='inner' alias='aa'>
                               <attribute name='mobilephone' />
                               <attribute name='emailaddress1' />
                               </link-entity>
                               </entity>
                          </fetch>";
                EntityCollection entColl = _CrmService.RetrieveMultiple(new FetchExpression(xmlquery));
                if (entColl.Entities.Count > 0)
                {
                    EntityReference entity = entColl.Entities[0].GetAttributeValue<EntityReference>("customerid");
                    CustomerID = entity.Id.ToString();
                    CustomerName = entity.Name.ToString();
                    Amount = decimal.Round(Convert.ToDecimal(entColl.Entities[0].GetAttributeValue<Money>("hil_receiptamount").Value.ToString()), 2);
                    AddressID = entColl.Entities[0].GetAttributeValue<EntityReference>("hil_serviceaddress").Id.ToString();
                    MobileNumber = entColl.Entities[0].GetAttributeValue<AliasedValue>("aa.mobilephone").Value.ToString();
                    Email = entColl.Entities[0].GetAttributeValue<AliasedValue>("aa.emailaddress1").Value.ToString();
                }
            }
            PaymentDetails objPayment = new PaymentDetails();
            Entity AddressCol = _CrmService.Retrieve("hil_address", new Guid(AddressID), new ColumnSet("hil_state", "hil_businessgeo", "hil_pincode", "hil_branch"));
            Entity entBranch = _CrmService.Retrieve("hil_branch", AddressCol.GetAttributeValue<EntityReference>("hil_branch").Id, new ColumnSet("hil_mamorandumcode"));
            string _mamorandumCode = "";
            string _txnid = "";
            if (entBranch.Attributes.Contains("hil_mamorandumcode"))
            {
                _mamorandumCode = entBranch.GetAttributeValue<string>("hil_mamorandumcode");
            }
            //try
            //{
            //    _txnid = getTransactionID(_CrmService, orderid);
            //}
            //catch (Exception ex)
            //{
            //    return (objPayment, new RequestStatus
            //    {
            //        StatusCode = (int)HttpStatusCode.BadRequest,
            //        Message = ex.Message
            //    });
            //}
            Entity paymentReceipt = new Entity("hil_paymentreceipt");
            //paymentReceipt["hil_transactionid"] = _txnid;
            paymentReceipt["hil_mobilenumber"] = MobileNumber;
            paymentReceipt["hil_email"] = Email;
            paymentReceipt["hil_orderid"] = new EntityReference("salesorder", orderid);
            paymentReceipt["hil_amount"] = new Money(decimal.Round(Convert.ToDecimal(Amount), 2));
            paymentReceipt["hil_memorandumcode"] = _mamorandumCode;
            paymentReceipt["hil_receiptdate"] = DateTime.Now.AddMinutes(330);
            paymentReceipt["hil_tokenexpireson"] = DateTime.Now.AddMinutes(330 + _Waittime);
            paymentReceipt["hil_paymentstatus"] = new OptionSetValue(1);
            Guid receiptId = _CrmService.Create(paymentReceipt);

            paymentReceipt = _CrmService.Retrieve(paymentReceipt.LogicalName, receiptId, new ColumnSet("hil_transactionid"));

            _txnid = paymentReceipt.GetAttributeValue<string>("hil_transactionid").ToString();// + Counter.ToString().PadLeft(3, '0');

            SendPayNowRequest req = new SendPayNowRequest();
            RequestData requestData = new RequestData();

            requestData.amount = Math.Round(Amount, 2).ToString();
            requestData.txnid = _txnid;
            requestData.firstname = CustomerName;
            requestData.email = Email;
            requestData.productinfo = _mamorandumCode;
            requestData.udf2 = "D365";
            requestData.udf1 = CustomerName;
            req.businessType = "B2C";
            req.paymentgateway_type = "PayUBiz";
            req.IM_PROJECT = "D365";
            req.RequestData = requestData;
            IntegrationConfiguration inconfig = CommonMethods.GetIntegrationConfiguration(_CrmService, _PayNowKey);

            var data = new StringContent(JsonConvert.SerializeObject(req), Encoding.UTF8, "application/json");

            RequestStatus objreq = new RequestStatus();
            HttpClient client = new HttpClient();
            var byteArray = Encoding.ASCII.GetBytes(inconfig.userName + ":" + inconfig.password);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
            HttpResponseMessage response = client.PostAsync(inconfig.url, data).Result;
            var obj = JsonConvert.DeserializeObject<PayNowPayu>(response.Content.ReadAsStringAsync().Result);
            if (obj.status == "Success")
            {
                //Entity reschudleOrder = new Entity("hil_paymentreceipt", receiptId);
                Entity paymentreceiptEnt = new Entity("hil_paymentreceipt", receiptId);
                paymentreceiptEnt["hil_hashname"] = obj.HashCode;
                paymentreceiptEnt["hil_salt"] = obj.salt;
                paymentreceiptEnt["hil_key"] = obj.key;
                _CrmService.Update(paymentreceiptEnt);

                //paymentreceiptEnt["hil_transactionid"] = _txnid;
                //paymentreceiptEnt["hil_mobilenumber"] = MobileNumber;
                //paymentreceiptEnt["hil_email"] = Email;
                //paymentreceiptEnt["hil_hashname"] = obj.HashCode;
                //paymentreceiptEnt["hil_salt"] = obj.salt;
                //paymentreceiptEnt["hil_key"] = obj.key;
                //paymentreceiptEnt["hil_orderid"] = new EntityReference("salesorder", orderid);
                //paymentreceiptEnt["hil_amount"] = new Money(decimal.Round(Convert.ToDecimal(Amount), 2));
                //paymentreceiptEnt["hil_memorandumcode"] = _mamorandumCode;
                //paymentreceiptEnt["hil_receiptdate"] = DateTime.Now;
                //paymentreceiptEnt["hil_paymentstatus"] = new OptionSetValue(1);
                //_CrmService.Update(paymentreceiptEnt);

                objPayment.TransactionID = _txnid;
                objPayment.HashName = obj.HashCode;
                objPayment.Salt = obj.salt;
                objPayment.Key = obj.key;
                objPayment.Amount = decimal.Round(Convert.ToDecimal(Amount), 2).ToString();
                objPayment.MamorandumCode = _mamorandumCode;
                objPayment.MobileNumber = MobileNumber;
                objPayment.Emailid = Email;
                objPayment.CustomerName = CustomerName;
                objPayment.Surl = _Surl;
                objPayment.Furl = _Furl;
                objreq.StatusCode = (int)HttpStatusCode.OK;
                objreq.Message = obj.message;
            }
            else
            {
                objreq.StatusCode = (int)HttpStatusCode.BadRequest;
                objreq.Message = obj.message;
            }
            return (objPayment, objreq);
        }
        private string getImgUrl(Guid CategoryId)
        {
            string xmlqery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' top='1' distinct='false'>
                  <entity name='hil_productcatalog'>
                    <attribute name='hil_productcatalogid' />
                    <attribute name='hil_name' />
                    <attribute name='hil_productimage' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                      <condition attribute='hil_productcode' operator='eq' value='{CategoryId}' />
                    </filter>
                  </entity>
                </fetch>";
            EntityCollection entitycol = _CrmService.RetrieveMultiple(new FetchExpression(xmlqery));
            if (entitycol.Entities.Count > 0)
                return entitycol.Entities[0].GetAttributeValue<string>("hil_productimage");
            else
                return "";
        }
        private async Task<bool> ValidateALaCartePriceAndDiscount(CreateOrder services)
        {
            int today = int.Parse(DateTime.Now.ToString("yyyyMMdd"));
            List<LTTABLE> lstALaCarteServices = new List<LTTABLE>();
            ALaCartePriceListResponse? responseData = null;
            bool matchFound = false;
            try
            {
                foreach (Services service in services.ServiceList)
                {
                    string ServiceName = _CrmService.Retrieve("product", new Guid(service.ServiceId), new ColumnSet("name")).GetAttributeValue<string>("name");
                    lstALaCarteServices.Add(new LTTABLE
                    {
                        MATNR = ServiceName,
                        KBETR = decimal.Parse(service.MRP).ToString("F2"),
                        KSCHL = "ZPR0"
                    });
                    if (services.PaymentType == "1")
                    {
                        lstALaCarteServices.Add(new LTTABLE
                        {
                            MATNR = ServiceName,
                            KBETR = decimal.Parse(service.DiscountPercent).ToString("F2"),
                            KSCHL = "ZDAM"
                        });
                    }
                }
                var requestData = new ALaCartePriceListRequest
                {
                    data = lstALaCarteServices.Select(service => new LTTABLE
                    {
                        MATNR = service.MATNR
                    }).ToList(),
                    IM_PROJECT = "D365"
                };

                var inconfig = CommonMethods.GetIntegrationConfiguration(_CrmService, "ValidateALaCartePrice");
                using (var client = new HttpClient())
                {
                    var byteArray = Encoding.ASCII.GetBytes($"{inconfig.userName}:{inconfig.password}");
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));

                    var data = new StringContent(JsonConvert.SerializeObject(requestData), Encoding.UTF8, "application/json");
                    var response = await client.PostAsync(inconfig.url, data);

                    var responseContent = await response.Content.ReadAsStringAsync();
                    _logger.LogToFile(new Exception(string.Format("CreateServiceOrder|Manager|1|{0}|ValidateALaCartePriceAndDiscount", "SAP API Response Content:" + responseContent)));
                    if (!string.IsNullOrWhiteSpace(responseContent))
                    {
                        responseData = JsonConvert.DeserializeObject<ALaCartePriceListResponse>(responseContent);
                    }
                    if (responseData?.LT_TABLE != null && responseData.LT_TABLE.Count > 0)
                    {
                        // var nonMatchingItems = new List<LTTABLE>();
                        foreach (LTTABLE service in lstALaCarteServices)
                        {
                            matchFound = responseData.LT_TABLE.Any(responseService =>
                              responseService.MATNR.Trim().Equals(service.MATNR, StringComparison.OrdinalIgnoreCase) &&
                              Math.Round(decimal.Parse(responseService.KBETR.Replace("-", "").Trim()), 0).ToString("F2").Equals(service.KBETR, StringComparison.OrdinalIgnoreCase) &&
                              responseService.KSCHL.Trim().Equals(service.KSCHL, StringComparison.OrdinalIgnoreCase) &&
                              responseService.DATBI >= today);

                            if (!matchFound && service.KSCHL == "ZDAM")
                            {
                                if (service.KBETR == "0.00")
                                {
                                    matchFound = true;
                                }
                            }
                            if (!matchFound)
                            {
                                return false;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception :" + ex.ToString());
            }
            finally
            {
                // Log method parameter
                _logger.LogToFile(new Exception(string.Format("CreateServiceOrder|Manager|2|{0}|ValidateALaCartePriceAndDiscount", "CreateServiceOrder Request:" + JsonConvert.SerializeObject(services))));
                // Log API response
                _logger.LogToFile(new Exception(string.Format("CreateServiceOrder|Manager|3|{0}|ValidateALaCartePriceAndDiscount", "ValidateALaCartePriceAndDiscount Result:" + matchFound.ToString())));
            }
            return matchFound;
        }
        public async Task<RequestStatus> RescheduleService(ReschuduleService objserviceReSec)
        {
            List<OrdersList> objServiceOrdersList = new List<OrdersList>();
            try
            {
                QueryExpression queryExpression = new QueryExpression("salesorder");
                queryExpression.ColumnSet = new ColumnSet("salesorderid", "requestdeliveryby", "hil_preferreddaytime");
                queryExpression.Criteria.AddCondition("salesorderid", ConditionOperator.Equal, objserviceReSec.orderId);
                EntityCollection orderEntCol = _CrmService.RetrieveMultiple(queryExpression);
                if (orderEntCol.Entities.Count > 0)
                {
                    Entity reschudleOrder = new Entity("salesorder", new Guid(objserviceReSec.orderId));
                    reschudleOrder["requestdeliveryby"] = Convert.ToDateTime(objserviceReSec.PreferredDate);
                    reschudleOrder["hil_preferreddaytime"] = new OptionSetValue(Convert.ToInt16(objserviceReSec.PreferredDateTime));
                    _CrmService.Update(reschudleOrder);
                    return (new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.OK,
                        Message = CommonMessage.SuccessMsg,
                    });
                }
                else
                {
                    return (new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = CommonMessage.InvalidOrderMsg,
                    });
                }
            }
            catch (Exception ex)
            {
                return (new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        public async Task<(List<OrdersList>, RequestStatus)> GetServiceRequestList(ServiceRequest objserviceReq)
        {
            List<OrdersList> objServiceOrdersList = new List<OrdersList>();
            try
            {
                string xmlquery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='salesorder'>
                                        <attribute name='name'/>
                                        <attribute name='customerid'/>
                                        <attribute name='totalamount'/>
                                        <attribute name='salesorderid'/>
                                        <attribute name='statuscode'/>
                                        <attribute name='hil_preferreddaytime'/>
                                        <attribute name='requestdeliveryby'/>
                                        <attribute name='hil_productdivision'/>
                                        <attribute name='hil_paymentstatus' />
                                        <attribute name='hil_modeofpayment' />
                                        <attribute name='ownerid'/>
                                        <order attribute='createdon' descending='true'/>
                                        <filter type='and'>
                                        <condition attribute='customerid' operator='eq'  value='{objserviceReq.CustomerId}'/>
                                        </filter>
                                        </entity>
                                        </fetch>";
                EntityCollection orderEntCol = _CrmService.RetrieveMultiple(new FetchExpression(xmlquery));
                if (orderEntCol.Entities.Count > 0)
                {
                    foreach (var item in orderEntCol.Entities)
                    {
                        OrdersList order = new OrdersList();
                        order.OrderGUID = item.Id.ToString();
                        order.OrderId = item.GetAttributeValue<string>("name");
                        order.ProductCategory = item.Contains("hil_productdivision") ? item.GetAttributeValue<EntityReference>("hil_productdivision").Name : "";
                        order.OrderAmount = item.Contains("totalamount") ? Convert.ToDouble(item.GetAttributeValue<Money>("totalamount").Value).ToString() : "";
                        order.OrderStatus = item.Contains("statuscode") ? item.FormattedValues["statuscode"].ToString() : "";
                        order.AssignTo = item.Contains("ownerid") ? item.GetAttributeValue<EntityReference>("ownerid").Name : "";

                        int OrderStatus = item.Contains("hil_paymentstatus") ? item.GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value : 0;
                        int ModeOfPayment = item.Contains("hil_modeofpayment") ? item.GetAttributeValue<OptionSetValue>("hil_modeofpayment").Value : 0;
                        if ((ModeOfPayment == 1 && OrderStatus == 2) || ModeOfPayment != 1) //Online & Success || Offline
                        {
                            QueryExpression query = new QueryExpression("salesorderdetail");
                            query.ColumnSet = new ColumnSet("productid", "priceperunit", "quantity");
                            query.Criteria.AddCondition("salesorderid", ConditionOperator.Equal, order.OrderGUID);
                            EntityCollection servicelineEntCol = _CrmService.RetrieveMultiple(query);
                            order.NumberOfServices = servicelineEntCol.Entities.Count.ToString();
                            order.EngMobileNo = _EngMobileNo;
                        }
                        objServiceOrdersList.Add(order);
                    }
                    return (objServiceOrdersList, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.OK,
                        Message = CommonMessage.SuccessMsg,
                    });
                }
                else
                {
                    return (objServiceOrdersList, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = CommonMessage.NoRecordExit,
                    });
                }
            }
            catch (Exception ex)
            {
                return (objServiceOrdersList, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        public async Task<(ServiceRequestData, RequestStatus)> GetServiceRequestDetails(OrderDetailRequest objservice)
        {
            ServiceRequestData _ServiceRequestDeatils = new ServiceRequestData();
            _ServiceRequestDeatils.ServiceJobDetails = new List<ServiceJobDetail>();
            try
            {
                string xmlquery = $@"<fetch version='1.0' output-format='xml-platform' top='1' mapping='logical' distinct='false'>
                                        <entity name='salesorder'>
                                        <attribute name='name'/>
                                        <attribute name='hil_productdivision'/>
                                        <attribute name='hil_serviceaddress'/>
                                        <attribute name='totalamount'/>
                                        <attribute name='salesorderid'/>
                                        <attribute name='hil_modeofpayment'/>
                                        <attribute name='hil_preferreddaytime'/>
                                        <attribute name='requestdeliveryby'/>
                                        <attribute name='hil_paymentstatus'/>
                                        <attribute name='ownerid'/>
                                        <order attribute='name' descending='false'/>
                                        <filter type='and'>
                                           <condition attribute='salesorderid' operator='eq'  value='{objservice.OrderGuid}'/>
                                        </filter>
                                        <link-entity name='hil_mediagallery' from='hil_salesorder' to='salesorderid' alias='mg' link-type='outer'>
                                            <attribute name='hil_mediatype' />
                                            <attribute name='hil_url' />
                                            <filter type='and'>
                                                <condition attribute='hil_mediatype' operator='eq' value='7dfe9cfa-d370-ef11-a671-7c1e5232c9eb' />
                                            </filter>
                                        </link-entity>
                                        </entity>
                                        </fetch>";
                EntityCollection orderEntCol = _CrmService.RetrieveMultiple(new FetchExpression(xmlquery));
                if (orderEntCol.Entities.Count > 0)
                {
                    Entity orderEnt = orderEntCol.Entities[0];
                    Entity EntAddress = _CrmService.Retrieve("hil_address", orderEnt.GetAttributeValue<EntityReference>("hil_serviceaddress").Id, new ColumnSet("hil_fulladdress"));

                    int PaymentStatus = orderEnt.Contains("hil_paymentstatus") ? orderEnt.GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value : 0;
                    int ModeOfPayment = orderEnt.Contains("hil_modeofpayment") ? orderEnt.GetAttributeValue<OptionSetValue>("hil_modeofpayment").Value : 0;

                    _ServiceRequestDeatils.ServiceLocation = EntAddress.Contains("hil_fulladdress") ? EntAddress.GetAttributeValue<string>("hil_fulladdress") : "";

                    _ServiceRequestDeatils.ProductCategory = orderEnt.Contains("hil_productdivision") ? orderEnt.GetAttributeValue<EntityReference>("hil_productdivision").Name : "";
                    if ((ModeOfPayment == 1 && PaymentStatus == 2) || ModeOfPayment != 1) //Online & Success || Offline
                    {
                        _ServiceRequestDeatils.ServiceDateTime = orderEnt.Contains("requestdeliveryby") ? orderEnt.GetAttributeValue<DateTime>("requestdeliveryby").Date.ToString("dd-MM-yyyy") : "";
                    }
                    _ServiceRequestDeatils.OrderId = orderEnt.Contains("name") ? orderEnt.GetAttributeValue<string>("name") : "";
                    _ServiceRequestDeatils.OrderAmount = orderEnt.Contains("totalamount") ? Convert.ToDouble(orderEnt.GetAttributeValue<Money>("totalamount").Value).ToString() : "";
                    _ServiceRequestDeatils.PaymentStatus = PaymentStatus == 0 ? "" : PaymentStatus.ToString();
                    _ServiceRequestDeatils.PaymentMode = ModeOfPayment == 0 ? "" : ModeOfPayment.ToString();
                    _ServiceRequestDeatils.DownloadInvoiceLink = orderEnt.Contains("mg.hil_url") ? orderEnt.GetAttributeValue<AliasedValue>("mg.hil_url").Value.ToString() : "";

                    string query = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='salesorderdetail'>
                                           <attribute name='productid'/>
                                           <attribute name='hil_job'/>
                                           <attribute name='priceperunit'/>
                                           <attribute name='quantity'/>
                                           <attribute name='manualdiscountamount'/>
                                           <attribute name='salesorderdetailid'/>
                                           <attribute name='ownerid'/>
                                           <order attribute='productid' descending='false'/>
                                        <filter type='and'>
                                            <condition attribute='salesorderid' operator='eq' value='{orderEnt.Id}'/>
                                        </filter>
                                        <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='hil_job' link-type='outer' alias='ak'>
                                            <attribute name='msdyn_substatus'/>
                                        </link-entity>
                                        </entity>
                                        </fetch>";
                    EntityCollection servicelineEntCol = _CrmService.RetrieveMultiple(new FetchExpression(query));
                    if (servicelineEntCol.Entities.Count > 0)
                    {
                        foreach (Entity item in servicelineEntCol.Entities)
                        {
                            Entity entProductCatalog = new Entity();
                            if (item.Contains("productid"))
                                entProductCatalog = CommonMethods.GetServiceDetails(_CrmService, item.GetAttributeValue<EntityReference>("productid").Id);
                            ServiceJobDetail serviceDetails = new ServiceJobDetail();
                            if ((ModeOfPayment == 1 && PaymentStatus == 2) || ModeOfPayment != 1) //Online & Success || Offline
                            {
                                serviceDetails.EngMobileNo = _EngMobileNo;
                            }
                            serviceDetails.ServiceName = entProductCatalog != null ? entProductCatalog.GetAttributeValue<string>("hil_name") : "";
                            serviceDetails.AssignTo = item.Contains("ownerid") ? item.GetAttributeValue<EntityReference>("ownerid").Name : "";
                            serviceDetails.ServiceAmount = Convert.ToDouble(item.Contains("priceperunit") ? item.GetAttributeValue<Money>("priceperunit").Value : 0);
                            serviceDetails.JobId = item.Contains("hil_job") ? item.GetAttributeValue<EntityReference>("hil_job").Id.ToString() : "";
                            serviceDetails.JobStatus = item.Contains("ak.msdyn_substatus") ? ((EntityReference)(item.GetAttributeValue<AliasedValue>("ak.msdyn_substatus").Value)).Name : "";

                            serviceDetails.JobStatusTracker = new List<JobStatusTracker>();
                            _ServiceRequestDeatils.ServiceJobDetails.Add(serviceDetails);
                            if (!string.IsNullOrWhiteSpace(serviceDetails.JobId))
                            {
                                QueryExpression Jobquery = new QueryExpression("hil_jobtracker");
                                Jobquery.ColumnSet = new ColumnSet("hil_jobid", "modifiedon", "hil_jobsubstatus");
                                Jobquery.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, serviceDetails.JobId);
                                EntityCollection jobEntCol = _CrmService.RetrieveMultiple(Jobquery);
                                if (jobEntCol.Entities.Count > 0)
                                {
                                    foreach (var jobitem in jobEntCol.Entities)
                                    {
                                        JobStatusTracker objJob = new JobStatusTracker();
                                        objJob.JobStatus = jobitem.Contains("hil_jobsubstatus") ? jobitem.GetAttributeValue<EntityReference>("hil_jobsubstatus").Name.ToString() : "";
                                        objJob.JobStatusDateTime = jobitem.Contains("modifiedon") ? jobitem.GetAttributeValue<DateTime>("modifiedon").AddMinutes(330).ToString("dd-MM-yyyy hh:mm:ss") : "";
                                        serviceDetails.JobStatusTracker.Add(objJob);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    return (_ServiceRequestDeatils, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.BadRequest,
                        Message = CommonMessage.NoRecordExit,
                    });

                }
            }
            catch (Exception ex)
            {
                return (_ServiceRequestDeatils, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
            return (_ServiceRequestDeatils, new RequestStatus()
            {
                StatusCode = (int)HttpStatusCode.OK,
                Message = CommonMessage.SuccessMsg,
            });
        }
    }
}
