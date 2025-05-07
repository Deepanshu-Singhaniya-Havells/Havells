using HavellsSync_Data.IManager;
using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Extensions.Configuration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Globalization;
using System.Net;
using System.Net.Http.Headers;
using System.Numerics;
using System.Reflection;
using System.Runtime.Intrinsics.Arm;
using System.Text;
using System.Text.RegularExpressions;

namespace HavellsSync_Data.Manager
{
    public class AMCManager : IAMCManager
    {
        private readonly ICrmService _crmService;
        private static Guid PriceLevelForFGsale = new Guid("c05e7837-8fc4-ed11-b597-6045bdac5e78");//AMC Ominichannnel
        private readonly IConfiguration _configuration;
        private readonly int _Waittime;
        private readonly string _Paymenturlkey;
        private readonly string _PayNowKey;
        public AMCManager(ICrmService crmService, IConfiguration configuration)
        {
            Check.Argument.IsNotNull(nameof(crmService), crmService);
            this._crmService = crmService;
            Check.Argument.IsNotNull(nameof(configuration), configuration);
            _configuration = configuration;
            _Waittime = Convert.ToInt32(configuration.GetSection(ConfigKeys.CRMSettings + ":" + ConfigKeys.Waittime).Value);
            _Paymenturlkey = configuration.GetSection(ConfigKeys.CRMSettings + ":" + ConfigKeys.Paymenturlkey).Value;
            _PayNowKey = configuration.GetSection(ConfigKeys.CRMSettings + ":" + ConfigKeys.PayNowKey).Value;
        }

        #region to GetData LoginUserId field not used
        public async Task<(WarrantyContentRes, RequestStatus)> GetWarrantyContent(string SourceType, string LoginUserId)
        {
            WarrantyContentRes warrantyContent = new WarrantyContentRes();
            EntityCollection entcoll;
            try
            {
                if (_crmService != null)
                {
                    if (SourceType != "6")
                    {
                        return (warrantyContent, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = CommonMessage.BadRequestMsg
                        });
                    }
                    //string TokenExpiresAt = Convert.ToString(CommonMethods.getExpiryTime(_crmService, SessionId));
                    string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='hil_productcatalog'>
                                    <attribute name='hil_productcatalogid' />
                                    <attribute name='hil_plantclink' />
                                    <attribute name='hil_productcode' />
                                    <order attribute='hil_productcode' descending='false' />
                                        <link-entity name='product' from='productid' to='hil_productcode' link-type='inner' alias='ag'>
                                            <filter type='and'>
                                                <condition attribute='hil_hierarchylevel' operator='eq' value='2' />
                                            </filter>
                                        </link-entity>
                                    </entity>
                                    </fetch>";
                    entcoll = _crmService.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entcoll.Entities.Count > 0)
                    {
                        List<AMCCatagory> lstAMCCatagory = new List<AMCCatagory>();
                        foreach (Entity ent in entcoll.Entities)
                        {
                            AMCCatagory objAMCCatagory = new AMCCatagory();
                            objAMCCatagory.CategoryId = ent.Id;
                            objAMCCatagory.CategoryName = ent.Contains("hil_productcode") ? ent.GetAttributeValue<EntityReference>("hil_productcode").Name : null;
                            objAMCCatagory.Icon = ent.Contains("hil_plantclink") ? ent.GetAttributeValue<string>("hil_plantclink") : null;
                            lstAMCCatagory.Add(objAMCCatagory);
                        }
                        warrantyContent.AMCCatagories = lstAMCCatagory;
                    }
                    _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='hil_warrantymediagallery'>
                                    <attribute name='hil_warrantymediagalleryid' />
                                    <attribute name='hil_name' />
                                    <attribute name='hil_content' />
                                    <attribute name='hil_imagepath' />
                                    <attribute name='hil_displayindex' />
                                    <attribute name='hil_category' />
                                        <order attribute='hil_category' descending='false' />
                                        <order attribute='hil_displayindex' descending='false' />
                                    </entity>
                                    </fetch>";
                    entcoll = _crmService.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entcoll.Entities.Count > 0)
                    {
                        List<WarrantyMediaContent> lstWarrantyMediaContent = new List<WarrantyMediaContent>();
                        List<WarrantyDiscountBanner> lstWarrantyDiscountBanner = new List<WarrantyDiscountBanner>();
                        foreach (Entity ent in entcoll.Entities)
                        {
                            int Category = ent.Contains("hil_category") ? ent.GetAttributeValue<OptionSetValue>("hil_category").Value : 0;
                            if (Category == 1) //Banner
                            {
                                WarrantyDiscountBanner objWarrantyDiscountBanner = new WarrantyDiscountBanner();
                                objWarrantyDiscountBanner.Index = ent.Contains("hil_displayindex") ? ent.GetAttributeValue<int>("hil_displayindex") : 0;
                                objWarrantyDiscountBanner.URL = ent.Contains("hil_imagepath") ? ent.GetAttributeValue<string>("hil_imagepath") : null;
                                lstWarrantyDiscountBanner.Add(objWarrantyDiscountBanner);
                            }
                            else if (Category == 2) // Category
                            {
                                WarrantyMediaContent objWarrantyMediaContent = new WarrantyMediaContent();
                                objWarrantyMediaContent.Icon = ent.Contains("hil_imagepath") ? ent.GetAttributeValue<string>("hil_imagepath") : null;
                                objWarrantyMediaContent.Content = ent.Contains("hil_content") ? ent.GetAttributeValue<string>("hil_content") : null;
                                lstWarrantyMediaContent.Add(objWarrantyMediaContent);
                            }
                        }
                        warrantyContent.WarrantyMediaContents = lstWarrantyMediaContent;
                        warrantyContent.WarrantyDiscountBanners = lstWarrantyDiscountBanner;
                    }
                    //warrantyContent.AccessToken = SessionId;
                    //warrantyContent.SourceType = SourceType;
                    //warrantyContent.TokenExpiresAt = TokenExpiresAt;
                    RequestStatus requestStatus = new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.OK,
                        Message = "Success",
                        // TokenExpiresAt = TokenExpiresAt
                    };
                    warrantyContent.StatusCode = requestStatus.StatusCode;
                    return (warrantyContent, requestStatus);
                }
                else
                {
                    return (warrantyContent, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.ServiceUnavailable,
                        Message = CommonMessage.ServiceUnavailableMsg
                    });
                }
            }
            catch (Exception ex)
            {
                return (warrantyContent, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }

        public async Task<(AMCPlanRes, RequestStatus)> GetAMCPlan(AMCPlanParam AMCPlanParam, string LoginUserId)
        {
            AMCPlanRes AMCPlanDtls = new AMCPlanRes();
            try
            {
                QueryExpression query;
                Guid ProductCategoryId = Guid.Empty;
                Guid Model = Guid.Empty;
                Guid ProductSubcategoryId = Guid.Empty;
                DateTime InvoiceDate = new DateTime(1900, 1, 1);
                int ProductAgeing = 0;

                List<AMCPlanInfo> lstAMCPlanInfo = new List<AMCPlanInfo>();
                if (_crmService != null)
                {
                    query = new QueryExpression("hil_integrationsource");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, AMCPlanParam.SourceType);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);//Active
                    EntityCollection sourceEntColl = _crmService.RetrieveMultiple(query);
                    if (sourceEntColl.Entities.Count == 0)
                    {
                        return (AMCPlanDtls, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = "Invalid Source Type."
                        });
                    }
                    Guid sourceId = sourceEntColl.Entities[0].Id;
                    Entity entityCustomerAsset = _crmService.Retrieve("msdyn_customerasset", new Guid(AMCPlanParam.CustomerAssestId), new ColumnSet("msdyn_product", "hil_productcategory", "hil_invoicedate", "hil_productsubcategory"));
                    if (entityCustomerAsset != null)
                    {
                        ProductCategoryId = entityCustomerAsset.Contains("hil_productcategory") ? entityCustomerAsset.GetAttributeValue<EntityReference>("hil_productcategory").Id : Guid.Empty;
                        Model = entityCustomerAsset.Contains("msdyn_product") ? entityCustomerAsset.GetAttributeValue<EntityReference>("msdyn_product").Id : Guid.Empty;
                        InvoiceDate = entityCustomerAsset.Contains("hil_invoicedate") ? entityCustomerAsset.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330) : new DateTime(1900, 1, 1);
                        ProductSubcategoryId = entityCustomerAsset.Contains("hil_productsubcategory") ? entityCustomerAsset.GetAttributeValue<EntityReference>("hil_productsubcategory").Id : Guid.Empty;
                    }
                    if (ProductCategoryId == Guid.Empty || Model == Guid.Empty || ProductSubcategoryId == Guid.Empty)
                    {
                        return (AMCPlanDtls, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = "Category/Sub-Category/Model is missing."
                        });
                    }
                    Guid AddressId = Guid.Empty;
                    if (!Guid.TryParse(AMCPlanParam.AddressId, out AddressId))
                    {
                        return (AMCPlanDtls, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = "Address is required."
                        });
                    }
                    Entity entCustomerAddress = _crmService.Retrieve("hil_address", AddressId, new ColumnSet("hil_state", "hil_salesoffice"));
                    Guid stateId = entCustomerAddress.Contains("hil_state") ? entCustomerAddress.GetAttributeValue<EntityReference>("hil_state").Id : Guid.Empty;
                    Guid salesofficeId = entCustomerAddress.Contains("hil_salesoffice") ? entCustomerAddress.GetAttributeValue<EntityReference>("hil_salesoffice").Id : Guid.Empty;

                    ProductAgeing = (int)DateTime.Now.Date.Subtract(InvoiceDate).TotalDays;

                    string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                    <entity name='hil_amcplansetup'>
                                    <attribute name='hil_amcplansetupid' />
                                    <attribute name='hil_amcplan' />
                                    <order attribute='hil_amcplan' descending='false' />
                                    <filter type='and'>
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    <condition attribute='hil_model' operator='eq' value='{Model}' />
                                    <filter type='or'>
                                    <condition attribute='hil_applicablesource' operator='eq' value='{sourceId}' />
                                    <condition attribute='hil_applicablesource' operator='null' />
                                    </filter>
                                    <filter type='or'>
                                    <condition attribute='hil_ageingstart' operator='null' />
                                    <filter type='and'>
                                    <condition attribute='hil_ageingstart' operator='le' value='{ProductAgeing}' />
                                    <condition attribute='hil_ageingend' operator='ge' value='{ProductAgeing}' />
                                    </filter>
                                    </filter>
                                    </filter>
                                    <link-entity name='product' from='productid' to='hil_model' link-type='inner' alias='be'>
                                    <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='bf'>
                                    <filter type='and'>
                                    <condition attribute='hil_eligibleforamc' operator='eq' value='1' />
                                    </filter>
                                    </link-entity>
                                    </link-entity>
                                    <link-entity name='product' from='productid' to='hil_amcplan' link-type='inner' alias='bg'>
                                    <filter type='and'>
                                    <condition attribute='hil_hierarchylevel' operator='eq' value='910590001' />
                                    </filter>
                                    <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='pc'>
                                    <attribute name='hil_name' />   
                                    <attribute name='createdon' />   
                                    <attribute name='hil_plantclink' />   
                                    <attribute name='hil_planperiod' />   
                                    <attribute name='hil_notcovered' />
                                    <attribute name='hil_coverage' />
                                    <attribute name='hil_amctandc' />
                                    </link-entity>
                                    <link-entity name='productpricelevel' from='productid' to='productid' link-type='inner' alias='pricelist'>
                                    <attribute name='amount' />
                                    <filter type='and'>
                                    <condition attribute='pricelevelid' operator='eq' value='{PriceLevelForFGsale}' />
                                    </filter>
                                    </link-entity>
                                    </link-entity>
                                    </entity>
                                    </fetch>";

                    EntityCollection entCollProduct = _crmService.RetrieveMultiple(new FetchExpression(fetchXML));

                    if (entCollProduct.Entities.Count > 0)
                    {
                        foreach (Entity entProduct in entCollProduct.Entities)
                        {
                            AMCPlanInfo objAMCPlanInfo = new AMCPlanInfo();

                            if (entProduct.Contains("hil_amcplan"))
                            {
                                objAMCPlanInfo.PlanId = entProduct.GetAttributeValue<EntityReference>("hil_amcplan").Id;
                                objAMCPlanInfo.DiscountPercent = GetDiscountValue(_crmService, Model, sourceId, ProductAgeing, ProductCategoryId, ProductSubcategoryId, objAMCPlanInfo.PlanId, stateId, salesofficeId).ToString();
                                objAMCPlanInfo.MRP = decimal.Round((entProduct.Contains("pricelist.amount") ? ((Money)entProduct.GetAttributeValue<AliasedValue>("pricelist.amount").Value).Value : 0), 2).ToString();

                                objAMCPlanInfo.Coverage = entProduct.Contains("pc.hil_coverage") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_coverage").Value.ToString() : "";
                                objAMCPlanInfo.NonCoverage = entProduct.Contains("pc.hil_notcovered") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_notcovered").Value.ToString() : "";
                                objAMCPlanInfo.PlanName = entProduct.Contains("pc.hil_name") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_name").Value.ToString() : "";
                                objAMCPlanInfo.PlanPeriod = entProduct.Contains("pc.hil_planperiod") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_planperiod").Value.ToString() : "";
                                objAMCPlanInfo.PlanTCLink = entProduct.Contains("pc.hil_plantclink") ? entProduct.GetAttributeValue<AliasedValue>("pc.hil_plantclink").Value.ToString() : "";
                            }
                            lstAMCPlanInfo.Add(objAMCPlanInfo);
                        }
                    }
                    AMCPlanDtls.ModelNumber = AMCPlanParam.ModelNumber;
                    AMCPlanDtls.AMCPlanInfo = lstAMCPlanInfo;
                    RequestStatus requestStatus = new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.OK,
                        Message = "Success",
                    };
                    AMCPlanDtls.StatusCode = requestStatus.StatusCode;
                    return (AMCPlanDtls, requestStatus);
                }
                else
                {
                    return (AMCPlanDtls, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.ServiceUnavailable,
                        Message = CommonMessage.ServiceUnavailableMsg
                    });
                }
            }
            catch (Exception ex)
            {
                return (AMCPlanDtls, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        public async Task<(PaymentStatusRes, RequestStatus)> GetStatus(PaymentStatusParam PaymentStatusParam, string LoginUserId)
        {
            PaymentStatusRes paymentStatusRes = new PaymentStatusRes();
            try
            {
                Guid InvoiceID = Guid.Empty;
                Guid PaymentstatusId = Guid.Empty;
                StatusRequest reqParm = new StatusRequest();
                string TxnId = string.Empty;
                if (_crmService != null)
                {
                    string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' top='1' distinct='false'>
                                      <entity name='hil_paymentreceipt'>
                                        <attribute name='hil_paymentreceiptid' />
                                        <attribute name='hil_transactionid' />
                                        <attribute name='createdon' />
                                        <attribute name='hil_receiptdate' />
                                        <attribute name='hil_paymenturl' />
                                        <attribute name='hil_paymentstatus' />
                                        <attribute name='hil_paymentmode' />
                                        <attribute name='hil_orderid' />
                                        <attribute name='hil_mobilenumber' />
                                        <attribute name='hil_memorandumcode' />
                                        <attribute name='hil_email' />
                                        <attribute name='hil_bankreferenceid' />
                                        <attribute name='hil_amount' />
                                        <order attribute='createdon' descending='true' />
                                        <filter type='and'>
                                          <condition attribute='hil_orderid' operator='eq' value='{PaymentStatusParam.InvoiceID}' />
                                          <condition attribute='statecode' operator='eq' value='0' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                    EntityCollection entpaymentstatus = _crmService.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (entpaymentstatus.Entities.Count == 0)
                    {
                        paymentStatusRes.StatusCode = (int)HttpStatusCode.NotFound;
                        return (paymentStatusRes, new RequestStatus
                        {
                            StatusCode = (int)HttpStatusCode.OK,
                            Message = CommonMessage.InvoiceNoNotFound
                        });
                    }
                    else
                    {
                        int paymentstatus = entpaymentstatus.Entities[0].Contains("hil_paymentstatus") ? entpaymentstatus.Entities[0].GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value : 1;
                        if (paymentstatus == 1 || paymentstatus == 3)//Payment Initiated || In Progress
                        {
                            TxnId = entpaymentstatus.Entities[0].GetAttributeValue<string>("hil_transactionid");
                            string Status = CommonMethods.getTransactionStatus(_crmService, entpaymentstatus.Entities[0].Id, TxnId, PaymentStatusParam.InvoiceID, _Paymenturlkey);

                            if (Status == "Success" || Status == "Failed" || Status == "Pending")
                            {
                                paymentStatusRes.PaymentStatus = Status;
                                paymentStatusRes.StatusCode = (int)HttpStatusCode.OK;
                                return (paymentStatusRes, new RequestStatus
                                {
                                    StatusCode = (int)HttpStatusCode.OK
                                });
                            }
                            else
                            {
                                return (paymentStatusRes, new RequestStatus()
                                {
                                    StatusCode = (int)HttpStatusCode.BadRequest,
                                    Message = CommonMessage.InternalServerErrorMsg + Status
                                });
                            }
                        }
                        else if (paymentstatus == 4)//Paid
                        {
                            paymentStatusRes.PaymentStatus = "Success";
                            paymentStatusRes.StatusCode = (int)HttpStatusCode.OK;
                            return (paymentStatusRes, new RequestStatus
                            {
                                StatusCode = (int)HttpStatusCode.OK
                            });
                        }
                        else
                        {
                            paymentStatusRes.PaymentStatus = "Failed";
                            paymentStatusRes.StatusCode = (int)HttpStatusCode.OK;
                            return (paymentStatusRes, new RequestStatus
                            {
                                StatusCode = (int)HttpStatusCode.OK
                            });
                        }
                    }
                }
                else
                {
                    return (paymentStatusRes, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.ServiceUnavailable,
                        Message = CommonMessage.ServiceUnavailableMsg
                    });
                }
            }
            catch (Exception ex)
            {
                return (paymentStatusRes, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }

        public async Task<(AMCOrdersListRes, RequestStatus)> GetAMCOrders(AMCOrdersParam AMCOrdersParam, string MobileNumber)
        {
            AMCOrdersListRes lstAMCOrdersRes = new AMCOrdersListRes();
            lstAMCOrdersRes.AMCOrders = new List<AMCOrders>();
            StatusRequest reqParm = new StatusRequest();
            string TxnId = string.Empty;
            Guid CustomerGuId = Guid.Empty;
            try
            {
                if (_crmService != null)
                {
                    try
                    {
                        CustomerGuId = new Guid(AMCOrdersParam.CustomerGuId);
                    }
                    catch (Exception)
                    {
                        return (lstAMCOrdersRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = CommonMessage.InvalidCustomerGuid
                        });
                    }
                    Entity entity = _crmService.Retrieve("contact", CustomerGuId, new ColumnSet(false));

                    if (entity == null)
                    {
                        return (lstAMCOrdersRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.NotFound,
                            Message = CommonMessage.FotFoundMsg
                        });
                    }
                    string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='salesorder'>
                                        <attribute name='name' />
                                        <attribute name='customerid' />
                                        <attribute name='statuscode' />
                                        <attribute name='totalamount' />
                                        <attribute name='salesorderid' />
                                        <attribute name='hil_paymentstatus' />
                                        <attribute name='createdon' />
                                        <order attribute='createdon' descending='true' />
                                        <filter type='and'>
                                          <condition attribute='customerid' operator='eq' value='{AMCOrdersParam.CustomerGuId}' />
                                          <condition attribute='hil_ordertype' operator='eq' value='{{1F9E3353-0769-EF11-A670-0022486E4ABB}}' />
                                          <condition attribute='statecode' operator='eq' value='0' />
                                        </filter>
                                      </entity>
                                    </fetch>";

                    EntityCollection entityColl = _crmService.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (entityColl.Entities.Count > 0)
                    {
                        foreach (Entity ent in entityColl.Entities)
                        {
                            AMCOrders amcOrdersRes = new AMCOrders();
                            amcOrdersRes.InvoiceId = ent.Id.ToString();
                            amcOrdersRes.InvoiceDate = ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString("dd/MM/yyyy");
                            amcOrdersRes.InvoiceValue = ent.Contains("totalamount") ? decimal.Round(ent.GetAttributeValue<Money>("totalamount").Value, 2).ToString() : null;
                            //amcOrdersRes.InvoiceDescription =ent.Contains("description") ? ent.GetAttributeValue<string>("description").ToString() : null;
                            int PaymentStatus = ent.Contains("hil_paymentstatus") ? ent.GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value : 0;
                            if (PaymentStatus == 1 || PaymentStatus == 3)
                            {
                                amcOrdersRes.PaymentStatus = "Pending";
                            }
                            else if (PaymentStatus == 2)
                            {
                                amcOrdersRes.PaymentStatus = "Success";
                            }
                            else
                            {
                                amcOrdersRes.PaymentStatus = "Failed";
                            }
                            //fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                            //            <entity name='hil_paymentreceipt'>
                            //            <attribute name='hil_paymentreceiptid' />
                            //            <attribute name='hil_transactionid' />
                            //            <attribute name='hil_paymentstatus' />
                            //            <attribute name='hil_tokenexpireson' />
                            //            <order attribute='createdon' descending='true' />
                            //            <filter type='and'>
                            //                <condition attribute='hil_orderid' operator='eq' value='{ent.Id}' />
                            //            </filter>
                            //            </entity>
                            //            </fetch>";
                            //EntityCollection _entColPayment = _crmService.RetrieveMultiple(new FetchExpression(fetchXml));
                            //if (_entColPayment.Entities.Count > 0)
                            //{
                            //    int paymentstatus = _entColPayment.Entities[0].Contains("hil_paymentstatus") ? _entColPayment.Entities[0].GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value : 1;

                            //    if (paymentstatus == 1 || paymentstatus == 3)//Payment Initiated || In Progress
                            //    {
                            //        TxnId = _entColPayment.Entities[0].GetAttributeValue<string>("hil_transactionid");
                            //        string Status = CommonMethods.getTransactionStatus(_crmService, _entColPayment.Entities[0].Id, TxnId, ent.Id, _Paymenturlkey);

                            //        if (Status == "Success" || Status == "Failed" || Status == "Pending")
                            //        {
                            //            amcOrdersRes.PaymentStatus = Status;
                            //        }
                            //        else
                            //        {
                            //            return (lstAMCOrdersRes, new RequestStatus()
                            //            {
                            //                StatusCode = (int)HttpStatusCode.BadRequest,
                            //                Message = CommonMessage.InternalServerErrorMsg + Status
                            //            });
                            //        }
                            //    }
                            //    else if (paymentstatus == 4)//Paid
                            //    {
                            //        amcOrdersRes.PaymentStatus = "Success";
                            //    }
                            //    else
                            //    {
                            //        amcOrdersRes.PaymentStatus = "Failed";
                            //    }
                            //}
                            lstAMCOrdersRes.AMCOrders.Add(amcOrdersRes);
                        }
                        lstAMCOrdersRes.StatusCode = (int)HttpStatusCode.OK;
                    }
                }
                else
                {
                    return (lstAMCOrdersRes, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.ServiceUnavailable,
                        Message = CommonMessage.ServiceUnavailableMsg
                    });
                }
                return (lstAMCOrdersRes, new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.OK
                });
            }
            catch (Exception ex)
            {
                return (lstAMCOrdersRes, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }

        #endregion

        #region LoginUserid Used to GetData
        public async Task<(AMCProductDeatilsList, RequestStatus)> GetRegisteredProductList(string SourceType, string MobileNumber)
        {
            AMCRegisterdProduct objRegisteredProducts;
            RequestStatus objReqStatus = new RequestStatus();
            AMCProductDeatilsList lstRegisteredProducts = new AMCProductDeatilsList();
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                if (_crmService != null)
                {
                    var CustomerGuid = CommonMethods.getCustomerGuid(_crmService, MobileNumber);
                    if (CustomerGuid == Guid.Empty)
                    {
                        objReqStatus.StatusCode = (int)HttpStatusCode.Forbidden;
                        objReqStatus.Message = CommonMessage.CustomerguidMsg;
                        return (lstRegisteredProducts, objReqStatus);
                    }
                    else if (string.IsNullOrEmpty(SourceType))
                    {
                        objReqStatus.StatusCode = (int)HttpStatusCode.Forbidden;
                        objReqStatus.Message = CommonMessage.ExtSourceTypeMsg;
                        return (lstRegisteredProducts, objReqStatus);
                    }
                    Query = new QueryExpression("msdyn_customerasset");
                    Query.ColumnSet = new ColumnSet("hil_warrantytilldate", "hil_warrantysubstatus", "hil_warrantystatus", "hil_modelname", "hil_product", "hil_retailerpincode", "hil_purchasedfrom", "hil_invoicevalue", "hil_invoiceno", "hil_invoicedate", "hil_invoiceavailable", "hil_batchnumber", "hil_pincode", "msdyn_customerassetid", "createdon", "msdyn_product", "msdyn_name", "hil_productsubcategorymapping", "hil_productsubcategory", "hil_productcategory", "hil_source");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, CustomerGuid);
                    Query.AddOrder("createdon", OrderType.Descending);

                    entcoll = _crmService.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        objReqStatus.StatusCode = (int)HttpStatusCode.Forbidden;
                        objReqStatus.Message = CommonMessage.CustomerNotExitMsg;
                        return (lstRegisteredProducts, objReqStatus);
                    }
                    else
                    {
                        lstRegisteredProducts.ProductList = new List<AMCRegisterdProduct>();
                        foreach (Entity ent in entcoll.Entities)
                        {
                            objRegisteredProducts = new AMCRegisterdProduct();
                            objRegisteredProducts.ProductGuid = ent.GetAttributeValue<Guid>("msdyn_customerassetid");

                            if (ent.Attributes.Contains("hil_productcategory"))
                            {
                                objRegisteredProducts.ProductCategory = ent.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                                objRegisteredProducts.ProductCategoryId = ent.GetAttributeValue<EntityReference>("hil_productcategory").Id;
                            }
                            if (ent.Attributes.Contains("msdyn_product"))
                            {
                                objRegisteredProducts.ModelName = ent.GetAttributeValue<EntityReference>("msdyn_product").Name;
                                objRegisteredProducts.ModelCode = ent.GetAttributeValue<EntityReference>("msdyn_product").Name;
                                objRegisteredProducts.ModelId = ent.GetAttributeValue<EntityReference>("msdyn_product").Id.ToString();
                            }
                            objRegisteredProducts.ProductWarranty = new List<AMCProductWarranty>();

                            objRegisteredProducts.ProductWarranty = GetWarrantyDetails(objRegisteredProducts.ProductGuid, _crmService);

                            if (ent.Attributes.Contains("hil_modelname"))
                            { objRegisteredProducts.ModelName = ent.GetAttributeValue<string>("hil_modelname"); }
                            if (ent.Attributes.Contains("hil_productsubcategory"))
                            {
                                objRegisteredProducts.ProductSubCategory = ent.GetAttributeValue<EntityReference>("hil_productsubcategory").Name;
                                objRegisteredProducts.ProductSubCategoryId = ent.GetAttributeValue<EntityReference>("hil_productsubcategory").Id;
                            }
                            if (ent.Attributes.Contains("msdyn_name"))
                            { objRegisteredProducts.SerialNumber = ent.GetAttributeValue<string>("msdyn_name"); }

                            if (ent.Attributes.Contains("hil_invoiceno"))
                            { objRegisteredProducts.InvoiceNumber = ent.GetAttributeValue<string>("hil_invoiceno"); }
                            if (ent.Attributes.Contains("hil_invoicedate"))
                            { objRegisteredProducts.InvoiceDate = ent.GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330).ToString("dd/MM/yyyy"); }
                            if (ent.Attributes.Contains("hil_invoicevalue"))
                            { objRegisteredProducts.InvoiceValue = decimal.Round((ent.GetAttributeValue<decimal>("hil_invoicevalue")), 2); }

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
                            lstRegisteredProducts.ProductList.Add(objRegisteredProducts);
                        }
                    }
                    objReqStatus.StatusCode = (int)HttpStatusCode.OK;
                    objReqStatus.Message = CommonMessage.SuccessMsg;
                    lstRegisteredProducts.StatusCode = objReqStatus.StatusCode;
                    return (lstRegisteredProducts, objReqStatus);
                }
                else
                {
                    objReqStatus.StatusCode = (int)HttpStatusCode.ServiceUnavailable;
                    objReqStatus.Message = CommonMessage.ServiceUnavailableMsg;
                    return (lstRegisteredProducts, objReqStatus);
                }
            }
            catch (Exception ex)
            {

                objReqStatus.StatusCode = (int)HttpStatusCode.BadRequest;
                objReqStatus.Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper();
                return (lstRegisteredProducts, objReqStatus);
            }
        }
        public async Task<(InitiatePaymentRes, RequestStatus)> InitiatePayment(InitiatePaymentParam IPParam, string MobileNumber)
        {
            InitiatePaymentRes initiatePaymentRes = new InitiatePaymentRes();
            string _fetchXML = string.Empty;
            EntityCollection _entCol = null;
            EntityReference _entRefConsumer = null;
            EntityReference _entRefBillingAddress = null;
            EntityReference _entRefAMCPlan = null;
            EntityReference _entRefAsset = null;
            EntityReference _entRefOrder = null;
            OptionSetValue _paymentStatus = null;
            string _transactionId = string.Empty;
            EntityReference _entRefPaymentTransId = null;
            Entity _entObj = null;
            Entity _entAssetObj = null;
            DateTime _tokenExpiryOn = new DateTime(1900, 1, 1);
            ConsumerProfile _consumerProfile = new ConsumerProfile()
            {
                mobileNumber = MobileNumber,
                emailId = "customercare@havells.com"
            };
            try
            {
                string _requestData = JsonConvert.SerializeObject(IPParam);
                if (_crmService != null)
                {
                    if (!Validate.IsValidMobileNumber(MobileNumber))
                    {
                        return (initiatePaymentRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = CommonMessage.MobileNumberMsg
                        });
                    }
                    if (IPParam.SourceType != "6")
                    {
                        string msg = CommonMessage.SourcetypeMsg;
                        if (!string.IsNullOrWhiteSpace(IPParam.SourceType))
                            msg = CommonMessage.InvalidSourceTypeMsg;
                        return (initiatePaymentRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = msg
                        });
                    }
                    if (!Validate.IsvalidGuid(IPParam.AMCPlanID))
                    {
                        string msg = CommonMessage.MandatoryAMCPlanID;
                        if (!string.IsNullOrWhiteSpace(IPParam.AMCPlanID))
                            msg = CommonMessage.InvalidAMCPlanID;
                        return (initiatePaymentRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = msg
                        });
                    }
                    else
                    {
                        _entObj = _crmService.Retrieve("product", new Guid(IPParam.AMCPlanID), new ColumnSet(false));
                        if (_entObj == null)
                        {
                            return (initiatePaymentRes, new RequestStatus()
                            {
                                StatusCode = (int)HttpStatusCode.BadRequest,
                                Message = CommonMessage.InvalidAMCPlanID
                            });
                        }
                        else
                            _entRefAMCPlan = _entObj.ToEntityReference();
                    }
                    if (!Validate.IsvalidGuid(IPParam.AssestId))
                    {
                        string msg = CommonMessage.MandatotyAssestIdMsg;
                        if (!string.IsNullOrWhiteSpace(IPParam.AssestId))
                            msg = CommonMessage.InvalidAssestIdMsg;
                        return (initiatePaymentRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = msg
                        });
                    }
                    else
                    {
                        _entAssetObj = _crmService.Retrieve("msdyn_customerasset", new Guid(IPParam.AssestId), new ColumnSet("hil_invoicedate", "hil_invoiceno", "hil_purchasedfrom", "hil_retailerpincode"));
                        if (_entAssetObj != null)
                        {
                            if (_entAssetObj.Contains("hil_invoicedate"))
                                IPParam.DOP = _entAssetObj.GetAttributeValue<DateTime>("hil_invoicedate").ToString("yyyy-MM-dd");
                            else
                                IPParam.DOP = "1900-01-01";

                            _entRefAsset = _entAssetObj.ToEntityReference();
                        }
                        else
                        {
                            return (initiatePaymentRes, new RequestStatus()
                            {
                                StatusCode = (int)HttpStatusCode.BadRequest,
                                Message = CommonMessage.InvalidAssestIdMsg
                            });
                        }
                    }
                    if (!Validate.IsvalidGuid(IPParam.AddressID))
                    {
                        string msg = CommonMessage.MandatotyAddressIdMsg;
                        if (!string.IsNullOrWhiteSpace(IPParam.AddressID))
                            msg = CommonMessage.Invalidaddress;
                        return (initiatePaymentRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = msg
                        });
                    }
                    else
                    {
                        _entObj = _crmService.Retrieve("hil_address", new Guid(IPParam.AddressID), new ColumnSet("hil_fulladdress", "hil_state", "hil_salesoffice", "hil_businessgeo", "hil_pincode", "hil_branch"));
                        if (_entObj == null)
                        {
                            return (initiatePaymentRes, new RequestStatus()
                            {
                                StatusCode = (int)HttpStatusCode.BadRequest,
                                Message = CommonMessage.Invalidaddress
                            });
                        }
                        else
                            _entRefBillingAddress = _entObj.ToEntityReference();
                        Entity entBranch = _crmService.Retrieve("hil_branch", _entObj.GetAttributeValue<EntityReference>("hil_branch").Id, new ColumnSet("hil_mamorandumcode"));
                        if (entBranch.Attributes.Contains("hil_mamorandumcode"))
                        {
                            _consumerProfile.branchMemorandumCode = entBranch.GetAttributeValue<string>("hil_mamorandumcode");
                        }
                        if (_entObj.Attributes.Contains("hil_state"))
                        {
                            _consumerProfile.state = _entObj.GetAttributeValue<EntityReference>("hil_state").Name.ToLower();
                        }
                        if (_entObj.Attributes.Contains("hil_pincode"))
                        {
                            _consumerProfile.pinCode = _entObj.GetAttributeValue<EntityReference>("hil_pincode").Name.ToLower();
                        }
                        if (_entObj.Attributes.Contains("hil_salesoffice"))
                        {
                            _consumerProfile.salesOffice = _entObj.GetAttributeValue<EntityReference>("hil_salesoffice").Name.ToLower();
                        }
                        if (_entObj.Attributes.Contains("hil_fulladdress"))
                        {
                            _consumerProfile.fullAddress = _entObj.GetAttributeValue<string>("hil_fulladdress");
                            if (_consumerProfile.fullAddress.Trim().Length > 99)
                                _consumerProfile.fullAddress = _consumerProfile.fullAddress.Substring(0, 99);
                        }
                    }
                    if (IPParam.Price < 1)
                    {
                        return (initiatePaymentRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = CommonMessage.MandatotyPriceMsg
                        });
                    }
                    QueryExpression query = new QueryExpression("hil_integrationsource");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, IPParam.SourceType);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);//Active
                    EntityCollection sourceEntColl = _crmService.RetrieveMultiple(query);
                    if (sourceEntColl.Entities.Count == 0)
                    {
                        return (initiatePaymentRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = "Invalid Source Type."
                        });
                    }
                    Guid CustomerGuid = CommonMethods.getCustomerGuid(_crmService, MobileNumber);
                    _entObj = _crmService.Retrieve("contact", CustomerGuid, new ColumnSet("emailaddress1", "fullname"));
                    if (_entObj == null)
                    {
                        return (initiatePaymentRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = "Customer not Found."
                        });
                    }
                    else
                    {
                        _entRefConsumer = _entObj.ToEntityReference();
                        if (_entObj.Contains("emailaddress1"))
                            _consumerProfile.emailId = _entObj.GetAttributeValue<string>("emailaddress1");
                        if (_entObj.Contains("fullname"))
                            _consumerProfile.fullName = _entObj.GetAttributeValue<string>("fullname");
                    }

                    #region Initialise Defaults
                    EntityReference _entRefOrderType = new EntityReference("hil_ordertype", new Guid("1f9e3353-0769-ef11-a670-0022486e4abb"));
                    EntityReference _entRefServiceAccount = new EntityReference("hil_ordertype", new Guid("d166ba69-65da-ec11-a7b5-6045bdad2a19"));
                    EntityReference _entRefPriceList = new EntityReference("pricelevel", new Guid("c05e7837-8fc4-ed11-b597-6045bdac5e78"));
                    OptionSetValue _modeOfPayment = new OptionSetValue(1);// Online:1 
                    OptionSetValue _orderType = new OptionSetValue(690970002);//Service-Maintenance Based
                    EntityReference _entRefSellingSource = sourceEntColl.Entities[0].ToEntityReference();
                    #endregion

                    #region Check If Order Exist
                    bool _paymentLinkToBeSend = false;
                    //Payment Status: {1:Initiated,2:Failed,3:In-progress,4-Paid}
                    //Order Payment Status: {2:Success}
                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                             <entity name='salesorder'>
                               <attribute name='salesorderid' />
                               <order attribute='createdon' descending='true' />
                               <filter type='and'>
                                 <condition attribute='customerid' operator='eq' value='{_entRefConsumer.Id}' />
                                 <condition attribute='hil_sellingsource' operator='eq' value='{_entRefSellingSource.Id}' />
                                 <condition attribute='hil_paymentstatus' operator='ne' value='2' />
                               </filter>
                               <link-entity name='salesorderdetail' from='salesorderid' to='salesorderid' link-type='inner' alias='ab'>
                                 <filter type='and'>
                                   <condition attribute='hil_product' operator='eq' value='{_entRefAMCPlan.Id}' />
                                   <condition attribute='hil_customerasset' operator='eq' value='{_entRefAsset.Id}' />
                                 </filter>
                               </link-entity>
                             </entity>
                           </fetch>";
                    _entCol = _crmService.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (_entCol.Entities.Count > 0)
                    {
                        _entRefOrder = _entCol.Entities[0].ToEntityReference();
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                                   <entity name='hil_paymentreceipt'>
                                   <attribute name='hil_paymentreceiptid' />
                                   <attribute name='hil_transactionid' />
                                   <attribute name='hil_paymentstatus' />
                                   <attribute name='hil_tokenexpireson' />
                                   <order attribute='createdon' descending='true' />
                                   <filter type='and'>
                                       <condition attribute='hil_orderid' operator='eq' value='{_entRefOrder.Id}' />
                                   </filter>
                                   </entity>
                                   </fetch>";
                        EntityCollection _entColPayment = _crmService.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (_entColPayment.Entities.Count > 0)
                        {
                            _paymentStatus = _entColPayment.Entities[0].GetAttributeValue<OptionSetValue>("hil_paymentstatus");
                            _transactionId = _entColPayment.Entities[0].GetAttributeValue<string>("hil_transactionid");
                            _entRefPaymentTransId = _entColPayment.Entities[0].ToEntityReference();
                            _tokenExpiryOn = _entColPayment.Entities[0].GetAttributeValue<DateTime>("hil_tokenexpireson").AddMinutes(330);
                            if (_paymentStatus.Value == 4)//PAID
                            {
                                return (initiatePaymentRes, new RequestStatus
                                {
                                    StatusCode = (int)HttpStatusCode.Forbidden,
                                    Message = "Payment is already recieved against given details."
                                });
                            }
                            else if (_paymentStatus.Value == 1 || _paymentStatus.Value == 3)
                            {
                                _paymentStatus = CommonMethods.GetPaymentStatus(_crmService, _transactionId, _entRefPaymentTransId.Id, _Paymenturlkey);
                                if (_paymentStatus != null)
                                {
                                    if (_paymentStatus.Value == 4)//PAID
                                    {
                                        return (initiatePaymentRes, new RequestStatus
                                        {
                                            StatusCode = (int)HttpStatusCode.Forbidden,
                                            Message = "Payment is already recieved against given details."
                                        });
                                    }
                                    else
                                    {
                                        if (_tokenExpiryOn >= DateTime.Now)
                                        {
                                            return (initiatePaymentRes, new RequestStatus
                                            {
                                                StatusCode = (int)HttpStatusCode.Forbidden,
                                                Message = "Payment link has been already sent. Please wait and try after 10 minutes."
                                            });
                                        }
                                        else
                                        {
                                            _paymentLinkToBeSend = true;
                                            //Place Holder to Create Payment Receipt and Call PayU API
                                        }
                                    }
                                }
                                else
                                {
                                    return (initiatePaymentRes, new RequestStatus
                                    {
                                        StatusCode = (int)HttpStatusCode.Forbidden,
                                        Message = "Payment is not initiated."
                                    });
                                }
                            }
                            else if (_paymentStatus.Value == 2)
                            {
                                _paymentLinkToBeSend = true;
                                //Create Payment Receipt and Call PayU API
                            }
                        }
                        else
                        {
                            //Create Payment Receipt and Call PayU API
                            _paymentLinkToBeSend = true;
                        }
                    }
                    else
                    {
                        #region Create Order and Order Line
                        Guid _orderGuId = Guid.Empty;
                        Entity _entOrder = new Entity("salesorder");
                        _entOrder["customerid"] = _entRefConsumer;
                        _entOrder["msdyn_psastatusreason"] = new OptionSetValue(192350000);//Draft
                        _entOrder["msdyn_ordertype"] = _orderType;
                        _entOrder["msdyn_account"] = _entRefServiceAccount;
                        _entOrder["hil_sellingsource"] = _entRefSellingSource;
                        _entOrder["hil_serviceaddress"] = _entRefBillingAddress;
                        _entOrder["hil_ordertype"] = _entRefOrderType;
                        _entOrder["pricelevelid"] = _entRefPriceList;
                        _entOrder["hil_receiptamount"] = new Money(Convert.ToDecimal(IPParam.Price));
                        _entOrder["hil_modeofpayment"] = new OptionSetValue(1);//Online
                        _entOrder["hil_sourcereferencecode"] = "";
                        try
                        {
                            _orderGuId = _crmService.Create(_entOrder);
                            _entRefOrder = new EntityReference("salesorder", _orderGuId);
                        }
                        catch (Exception ex)
                        {
                            return (initiatePaymentRes, new RequestStatus
                            {
                                StatusCode = (int)HttpStatusCode.Forbidden,
                                Message = "ERROR: Something went wrong."
                            });
                        }
                        Entity _entOrderLine = new Entity("salesorderdetail");
                        _entOrderLine["salesorderid"] = new EntityReference("salesorder", _orderGuId);
                        _entOrderLine["productid"] = _entRefAMCPlan;
                        _entOrderLine["hil_product"] = _entRefAMCPlan;
                        _entOrderLine["quantity"] = new decimal(1);
                        //_entOrderLine["priceperunit"] = new Money(Convert.ToDecimal(line.PricePerUnit));
                        //_entOrderLine["baseamount"] = new Money(Convert.ToDecimal(line.Amount));
                        _entOrderLine["uomid"] = new EntityReference("uom", new Guid("0359d51b-d7cf-43b1-87f6-fc13a2c1dec8"));
                        // _entOrderLine["ownerid"] = new EntityReference("systemuser", new Guid(Technicianid));
                        _entOrderLine["hil_customerasset"] = _entRefAsset;
                        if (_entAssetObj.Contains("hil_invoicedate"))
                            _entOrderLine["hil_invoicedate"] = _entAssetObj.GetAttributeValue<DateTime>("hil_invoicedate").AddHours(330);
                        else
                            _entOrderLine["hil_invoicedate"] = new DateTime(1900, 1, 1);
                        _entOrderLine["hil_invoicenumber"] = _entAssetObj.Contains("hil_invoiceno") ? _entAssetObj.GetAttributeValue<string>("hil_invoiceno") : null;
                        _entOrderLine["hil_invoicevalue"] = new Money(_entAssetObj.GetAttributeValue<decimal>("hil_invoicevalue"));
                        _entOrderLine["hil_purchasefrom"] = _entAssetObj.Contains("hil_purchasedfrom") ? _entAssetObj.GetAttributeValue<string>("hil_purchasedfrom") : null;
                        _entOrderLine["hil_purchasefromlocation"] = _entAssetObj.Contains("hil_retailerpincode") ? _entAssetObj.GetAttributeValue<string>("hil_retailerpincode") : null;
                        try
                        {
                            _crmService.Create(_entOrderLine);
                        }
                        catch (Exception ex)
                        {
                            return (initiatePaymentRes, new RequestStatus
                            {
                                StatusCode = (int)HttpStatusCode.Forbidden,
                                Message = "ERROR: Something went wrong."
                            });
                        }
                        #endregion
                        _paymentLinkToBeSend = true;
                        //Create Payment Receipt and Call PayU API
                    }
                    string ResultMessage = "";
                    if (_paymentLinkToBeSend)
                    {
                        ResInvoiceInfo resInvoiceInfo = PayNow(_entRefOrder.Id, _entRefAMCPlan.Id, IPParam.Price, _consumerProfile, _crmService);
                        if (!resInvoiceInfo.ResultStatus)
                        {
                            return (initiatePaymentRes, new RequestStatus
                            {
                                StatusCode = (int)HttpStatusCode.Forbidden,
                                Message = resInvoiceInfo.ResultMessage
                            });
                        }
                        ResultMessage = resInvoiceInfo.ResultMessage;
                        initiatePaymentRes.InvoiceID = resInvoiceInfo.InvoiceID.ToString();
                        initiatePaymentRes.MamorandumCode = _consumerProfile.branchMemorandumCode;
                        initiatePaymentRes.TransactionID = resInvoiceInfo.TransactionID;
                        initiatePaymentRes.HashName = resInvoiceInfo.HashCode;
                        initiatePaymentRes.MobileNumber = MobileNumber;
                        initiatePaymentRes.Emailid = _consumerProfile.emailId;
                        initiatePaymentRes.UserName = _consumerProfile.fullName;
                        initiatePaymentRes.Surl = _configuration.GetSection(ConfigKeys.CRMSettings + ":" + ConfigKeys.Surl).Value;
                        initiatePaymentRes.Furl = _configuration.GetSection(ConfigKeys.CRMSettings + ":" + ConfigKeys.Furl).Value;
                        initiatePaymentRes.Key = resInvoiceInfo.key;
                        initiatePaymentRes.Salt = resInvoiceInfo.salt;
                        initiatePaymentRes.StatusCode = (int)HttpStatusCode.OK;
                    }
                    return (initiatePaymentRes, new RequestStatus
                    {
                        StatusCode = (int)HttpStatusCode.OK,
                        Message = ResultMessage
                    });
                    #endregion
                }
                else
                {
                    return (initiatePaymentRes, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.ServiceUnavailable,
                        Message = CommonMessage.ServiceUnavailableMsg
                    });
                }
            }
            catch (Exception ex)
            {
                return (initiatePaymentRes, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        public async Task<(List<TranscationHistory>, RequestStatus)> GetTransactionDetails_Old(AMCOrdersParam AMCOrdersParam, string MobileNumber)
        {
            List<TranscationHistory> transcationHistories = new List<TranscationHistory>();
            List<TranscationHistoryWithOrder> tranHisList = new List<TranscationHistoryWithOrder>();
            try
            {
                if (_crmService != null)
                {
                    string fetchXml = $@"<fetch>
                                <entity name='invoice'>
                                    <attribute name='invoiceid' />
                                    <attribute name='customerid' />
                                    <attribute name='customeridname' />
                                    <attribute name='hil_productcode' />
                                    <attribute name='msdyn_invoicedate' />
                                    <attribute name='invoicenumber' />
                                    <attribute name='totalamount' />
                                    <attribute name='hil_receiptamount' />
                                    <attribute name='hil_modelcode' />
                                    <attribute name='hil_newserialnumber' />
                                    <order attribute='createdon' descending='true' />
                                    <filter type='and'>
                                        <condition attribute='customerid' operator='eq' value='{AMCOrdersParam.CustomerGuId}' />
                                    </filter>
                                    <link-entity name='product' from='productid' to='hil_productcode' link-type='inner' alias='ag'>
                                        <filter type='and'>
                                        <condition attribute='hil_hierarchylevel' operator='eq' value='910590001' /> 
                                        </filter>
                                     </link-entity>
                                </entity>
                            </fetch>";
                    EntityCollection entityColl = _crmService.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (entityColl.Entities.Count > 0)
                    {
                        foreach (Entity ent in entityColl.Entities)
                        {
                            Guid hil_productcode = ent.Contains("hil_productcode") ? ent.GetAttributeValue<EntityReference>("hil_productcode").Id : Guid.Empty;
                            var ModelCode = ent.Contains("hil_modelcode") ? ent.GetAttributeValue<EntityReference>("hil_modelcode").Name : string.Empty; //ModelNumber(or ModelCode) 
                            var receiptamount = Math.Round(ent.Contains("hil_receiptamount") ? ent.GetAttributeValue<Money>("hil_receiptamount").Value : 0, 2).ToString();
                            if (string.IsNullOrEmpty(ModelCode))
                            {
                                var serialnumber = ent.Contains("hil_newserialnumber") ? ent.GetAttributeValue<string>("hil_newserialnumber") : string.Empty;
                                QueryExpression modelquery = new QueryExpression("msdyn_customerasset");
                                modelquery.ColumnSet = new ColumnSet("hil_modelname", "hil_product", "msdyn_product", "hil_productcategory");
                                modelquery.Criteria = new FilterExpression(LogicalOperator.And);
                                modelquery.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, serialnumber);
                                Entity ModelDetailCollection = _crmService.RetrieveMultiple(modelquery).Entities[0];
                                ModelCode = ModelDetailCollection.Contains("msdyn_product") ? ModelDetailCollection.GetAttributeValue<EntityReference>("msdyn_product").Name : "";
                            }

                            QueryExpression AMCPlanquery = new QueryExpression("hil_productcatalog");
                            AMCPlanquery.ColumnSet = new ColumnSet("hil_productcatalogid", "hil_name", "hil_planperiod");
                            AMCPlanquery.Criteria = new FilterExpression(LogicalOperator.And);
                            AMCPlanquery.Criteria.AddCondition("hil_productcode", ConditionOperator.Equal, hil_productcode);
                            Entity EntityAMCPlan = _crmService.RetrieveMultiple(AMCPlanquery).Entities[0];

                            QueryExpression query = new QueryExpression("msdyn_paymentdetail");
                            query.ColumnSet = new ColumnSet("statuscode", "createdon", "msdyn_invoice", "msdyn_name", "msdyn_paymentamount");
                            query.Criteria.AddCondition("msdyn_invoice", ConditionOperator.Equal, ent.Id);
                            //query.Criteria.AddCondition("statuscode", ConditionOperator.NotEqual, 910590000);//Not Received
                            query.AddOrder("createdon", OrderType.Descending);
                            EntityCollection entpaymentstatus = _crmService.RetrieveMultiple(query);

                            if (entpaymentstatus.Entities.Count > 0)
                            {
                                var planperiod = EntityAMCPlan.Contains("hil_planperiod") ? EntityAMCPlan.GetAttributeValue<string>("hil_planperiod") : null;
                                var PlanName = EntityAMCPlan.Contains("hil_name") ? EntityAMCPlan.GetAttributeValue<string>("hil_name") : "";
                                foreach (Entity PaymetHistroyItem in entpaymentstatus.Entities)
                                {
                                    TranscationHistoryWithOrder objTranItem = new TranscationHistoryWithOrder();
                                    objTranItem.Transactionid = PaymetHistroyItem.Contains("msdyn_name") ? PaymetHistroyItem.GetAttributeValue<string>("msdyn_name") : "";    // PaymetHistroyItem.GetAttributeValue<EntityReference>("msdyn_paymentdetail").Name;
                                    objTranItem.TransactionDate = PaymetHistroyItem.Contains("createdon") ? PaymetHistroyItem.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() : null;  //ent.GetAttributeValue<DateTime>("msdyn_invoicedate").ToString("dd/MM/yyyy")
                                    objTranItem.TransactionOrder = PaymetHistroyItem.Contains("createdon") ? PaymetHistroyItem.GetAttributeValue<DateTime>("createdon").AddMinutes(330) : null;  //ent.GetAttributeValue<DateTime>("msdyn_invoicedate").ToString("dd/MM/yyyy")
                                    objTranItem.InvoiceId = PaymetHistroyItem.Contains("msdyn_invoice") ? PaymetHistroyItem.GetAttributeValue<EntityReference>("msdyn_invoice").Name : null;
                                    string Status = PaymetHistroyItem.FormattedValues["statuscode"].ToString();
                                    objTranItem.PlanDuration = planperiod; //EntityAMCPlan.Contains("hil_planperiod") ? EntityAMCPlan.GetAttributeValue<string>("hil_planperiod") : null;
                                    objTranItem.PlanName = PlanName;       //EntityAMCPlan.Contains("hil_name") ? EntityAMCPlan.GetAttributeValue<string>("hil_name") : "";
                                    objTranItem.Amount = receiptamount;   //Math.Round(ent.Contains("hil_receiptamount") ? ent.GetAttributeValue<Money>("hil_receiptamount").Value : 0, 2).ToString();
                                    objTranItem.ProductName = ModelCode;
                                    string PaymentStatus = "";
                                    // string PaymentStatus = (Status == "Received" ? "Success" : (Status == "Initiated"|| Status=="Active") ? "Pending" : Status);
                                    if (Status == "Initiated" || Status == "Active")
                                    {
                                        Status = CommonMethods.getTransactionStatus_Old(_crmService, objTranItem.Transactionid, ent.Id, _Paymenturlkey);
                                        PaymentStatus = (Status == "success" ? "Success" : (Status == "failure" ? "Failed" : "Pending"));
                                    }
                                    else
                                    {
                                        PaymentStatus = (Status == "Received" ? "Success" : (Status == "Initiated" || Status == "Active") ? "Pending" : Status);
                                    }
                                    if (PaymentStatus == "Pending")
                                    {
                                        objTranItem.InfoMessage = CommonMessage.PendingMsg;
                                    }
                                    if (PaymentStatus == "Failed")
                                    {
                                        objTranItem.InfoMessage = CommonMessage.FailedMsg;
                                    }
                                    if (PaymentStatus == "Success")
                                    {
                                        objTranItem.InfoMessage = CommonMessage.SuccessMsg;
                                    }
                                    objTranItem.PaymentStatus = PaymentStatus;

                                    tranHisList.Add(objTranItem);
                                }
                            }
                        }
                    }
                    List<TranscationHistoryWithOrder> lstSalesOrder = GetSalesOrderTransactions(_crmService, new Guid(AMCOrdersParam.CustomerGuId));
                    tranHisList.AddRange(lstSalesOrder);
                    transcationHistories = tranHisList.OrderByDescending(m => m.TransactionOrder).Select(m => new TranscationHistory
                    {
                        InvoiceId = m.InvoiceId,
                        Transactionid = m.Transactionid,
                        PlanName = m.PlanName,
                        ProductName = m.ProductName,
                        Amount = m.Amount,
                        PlanDuration = m.PlanDuration,
                        PaymentStatus = m.PaymentStatus,
                        TransactionDate = m.TransactionDate,
                        InfoMessage = m.InfoMessage
                    }).ToList();
                }
                else
                {
                    return (transcationHistories, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.ServiceUnavailable,
                        Message = CommonMessage.ServiceUnavailableMsg
                    });
                }
                return (transcationHistories, new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.OK
                });
            }
            catch (Exception ex)
            {
                return (transcationHistories, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }
        public async Task<(List<TranscationHistory>, RequestStatus)> GetTransactionDetails(AMCOrdersParam AMCOrdersParam, string MobileNumber)
        {
            List<TranscationHistory> lstTranscationHistory = new List<TranscationHistory>();
            try
            {
                if (_crmService != null)
                {
                    try
                    {
                        new Guid(AMCOrdersParam.CustomerGuId);
                    }
                    catch (Exception)
                    {
                        return (lstTranscationHistory, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = CommonMessage.InvalidCustomerGuid
                        });
                    }
                    string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='salesorder'>
                                        <attribute name='salesorderid'/>
                                        <attribute name='name'/>
                                        <attribute name='customerid'/>
                                        <attribute name='totalamount'/>
                                        <attribute name='statuscode'/>
                                        <attribute name='hil_paymentstatus'/>
                                        <attribute name='hil_productdivision'/>
                                        <attribute name='ownerid'/>
                                        <order attribute='createdon' descending='true' />
                                        <filter type='and'>
                                            <condition attribute='customerid' operator='eq'  value='{AMCOrdersParam.CustomerGuId}'/>                                           
                                        </filter>
                                        </entity>
                                        </fetch>";
                    EntityCollection entityColl = _crmService.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (entityColl.Entities.Count > 0)
                    {
                        foreach (Entity ent in entityColl.Entities)
                        {
                            TranscationHistory transcationHistory = new TranscationHistory();
                            transcationHistory.InvoiceId = ent.GetAttributeValue<string>("name");
                            transcationHistory.PlanName = ent.Contains("hil_productdivision") ? ent.GetAttributeValue<EntityReference>("hil_productdivision").Name : "";
                            transcationHistory.Amount = Math.Round(ent.Contains("totalamount") ? ent.GetAttributeValue<Money>("totalamount").Value : 0, 2).ToString();

                            fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' top='1' distinct='false'>
                                                  <entity name='hil_paymentreceipt'>
                                                    <attribute name='hil_paymentreceiptid' />
                                                    <attribute name='hil_transactionid' />
                                                    <attribute name='hil_paymentstatus' />
                                                    <attribute name='hil_bankreferenceid' />                                                    
                                                    <attribute name='createdon' />
                                                    <order attribute='createdon' descending='true' />
                                                    <filter type='and'>
                                                        <condition attribute='statecode' operator='eq' value='0' />
                                                        <condition attribute='hil_orderid' operator='eq' value='{ent.Id}' />
                                                    </filter>
                                                  </entity>
                                                </fetch>";
                            EntityCollection entPaymentreceipt = _crmService.RetrieveMultiple(new FetchExpression(fetchXML));
                            if (entPaymentreceipt.Entities.Count > 0)
                            {
                                int Status = entPaymentreceipt.Entities[0].GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value;
                                if (Status == 1 || Status == 3)
                                {
                                    transcationHistory.PaymentStatus = "Pending";
                                    transcationHistory.InfoMessage = CommonMessage.PendingMsg;
                                }
                                if (Status == 2 || Status == 5)
                                {
                                    transcationHistory.PaymentStatus = "Failed";
                                    transcationHistory.InfoMessage = CommonMessage.FailedMsg;
                                }
                                if (Status == 4)
                                {
                                    transcationHistory.PaymentStatus = "Success";
                                    transcationHistory.InfoMessage = CommonMessage.SuccessMsg;
                                }
                                transcationHistory.Transactionid = entPaymentreceipt.Entities[0].Contains("hil_transactionid") ? entPaymentreceipt.Entities[0].GetAttributeValue<string>("hil_transactionid") : "";
                                transcationHistory.TransactionDate = entPaymentreceipt.Entities[0].Contains("createdon") ? entPaymentreceipt.Entities[0].GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() : null;
                            }
                            string serviceName = "";
                            string query = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='salesorderdetail'>
                                            <attribute name='productid'/>
                                            <attribute name='salesorderdetailid'/>
                                            <order attribute='productid' descending='false'/>
                                            <filter type='and'>
                                                <condition attribute='salesorderid' operator='eq' value='{ent.Id}'/>
                                            </filter>
                                            </entity>
                                            </fetch>";
                            EntityCollection servicelineEntCol = _crmService.RetrieveMultiple(new FetchExpression(query));
                            if (servicelineEntCol.Entities.Count > 0)
                            {
                                foreach (Entity item in servicelineEntCol.Entities)
                                {
                                    Entity entProductCatalog = new Entity();
                                    if (item.Contains("productid"))
                                    {
                                        entProductCatalog = CommonMethods.GetServiceDetails(_crmService, item.GetAttributeValue<EntityReference>("productid").Id);
                                        string tempServiceName = entProductCatalog != null ? entProductCatalog.GetAttributeValue<string>("hil_name") : "";
                                        if (tempServiceName != "")
                                        {
                                            serviceName = serviceName + tempServiceName + ", ";
                                        }
                                    }
                                }
                                if (serviceName != "")
                                    serviceName = serviceName.Substring(0, serviceName.Length - 2);
                            }
                            transcationHistory.PlanDuration = servicelineEntCol.Entities.Count.ToString() + " Services";
                            transcationHistory.ProductName = serviceName;
                            lstTranscationHistory.Add(transcationHistory);
                        }
                    }
                }
                else
                {
                    return (lstTranscationHistory, new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.ServiceUnavailable,
                        Message = CommonMessage.ServiceUnavailableMsg
                    });
                }
                return (lstTranscationHistory, new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.OK
                });
            }
            catch (Exception ex)
            {
                return (lstTranscationHistory, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = CommonMessage.InternalServerErrorMsg + ex.Message.ToUpper()
                });
            }
        }

        #endregion

        #region Internaly issued Methods To GetValues
        private static List<TranscationHistoryWithOrder> GetSalesOrderTransactions(ICrmService service, Guid CustomerGuid)
        {
            List<TranscationHistoryWithOrder> lstTranscationHistories = new List<TranscationHistoryWithOrder>();
            try
            {
                string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='salesorder'>
                                        <attribute name='salesorderid'/>
                                        <attribute name='name'/>
                                        <attribute name='customerid'/>
                                        <attribute name='totalamount'/>
                                        <attribute name='statuscode'/>
                                        <attribute name='hil_paymentstatus'/>
                                        <attribute name='hil_productdivision'/>
                                        <attribute name='ownerid'/>
                                        <order attribute='createdon' descending='true' />
                                        <filter type='and'>
                                            <condition attribute='customerid' operator='eq'  value='{CustomerGuid}'/>                                           
                                        </filter>
                                        </entity>
                                        </fetch>";
                EntityCollection entityColl = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (entityColl.Entities.Count > 0)
                {
                    foreach (Entity ent in entityColl.Entities)
                    {
                        TranscationHistoryWithOrder transcationHistory = new TranscationHistoryWithOrder();
                        transcationHistory.InvoiceId = ent.GetAttributeValue<string>("name");
                        transcationHistory.PlanName = ent.Contains("hil_productdivision") ? ent.GetAttributeValue<EntityReference>("hil_productdivision").Name : "";
                        transcationHistory.Amount = Math.Round(ent.Contains("totalamount") ? ent.GetAttributeValue<Money>("totalamount").Value : 0, 2).ToString();

                        fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                  <entity name='hil_paymentreceipt'>
                                                    <attribute name='hil_paymentreceiptid' />
                                                    <attribute name='hil_transactionid' />
                                                    <attribute name='hil_paymentstatus' />
                                                    <attribute name='hil_bankreferenceid' />                                                    
                                                    <attribute name='createdon' />
                                                    <order attribute='createdon' descending='true' />
                                                    <filter type='and'>
                                                        <condition attribute='statecode' operator='eq' value='0' />
                                                        <condition attribute='hil_orderid' operator='eq' value='{ent.Id}' />
                                                    </filter>
                                                  </entity>
                                                </fetch>";
                        EntityCollection entPaymentreceipt = service.RetrieveMultiple(new FetchExpression(fetchXML));
                        if (entPaymentreceipt.Entities.Count > 0)
                        {
                            int Status = entPaymentreceipt.Entities[0].GetAttributeValue<OptionSetValue>("hil_paymentstatus").Value;
                            if (Status == 1 || Status == 3)
                            {
                                transcationHistory.PaymentStatus = "Pending";
                                transcationHistory.InfoMessage = CommonMessage.PendingMsg;
                            }
                            if (Status == 2 || Status == 5)
                            {
                                transcationHistory.PaymentStatus = "Failed";
                                transcationHistory.InfoMessage = CommonMessage.FailedMsg;
                            }
                            if (Status == 4)
                            {
                                transcationHistory.PaymentStatus = "Success";
                                transcationHistory.InfoMessage = CommonMessage.SuccessMsg;
                            }
                            transcationHistory.Transactionid = entPaymentreceipt.Entities[0].Contains("hil_transactionid") ? entPaymentreceipt.Entities[0].GetAttributeValue<string>("hil_transactionid") : "";
                            transcationHistory.TransactionDate = entPaymentreceipt.Entities[0].Contains("createdon") ? entPaymentreceipt.Entities[0].GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() : null;
                            transcationHistory.TransactionOrder = entPaymentreceipt.Entities[0].Contains("createdon") ? entPaymentreceipt.Entities[0].GetAttributeValue<DateTime>("createdon").AddMinutes(330) : null;
                        }
                        string serviceName = "";
                        string query = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='salesorderdetail'>
                                        <attribute name='productid'/>
                                        <attribute name='salesorderdetailid'/>
                                        <order attribute='productid' descending='false'/>
                                        <filter type='and'>
                                            <condition attribute='salesorderid' operator='eq' value='{ent.Id}'/>
                                        </filter>
                                        </entity>
                                        </fetch>";
                        EntityCollection servicelineEntCol = service.RetrieveMultiple(new FetchExpression(query));
                        if (servicelineEntCol.Entities.Count > 0)
                        {
                            foreach (Entity item in servicelineEntCol.Entities)
                            {
                                Entity entProductCatalog = new Entity();
                                if (item.Contains("productid"))
                                {
                                    entProductCatalog = CommonMethods.GetServiceDetails(service, item.GetAttributeValue<EntityReference>("productid").Id);
                                    string tempServiceName = entProductCatalog != null ? entProductCatalog.GetAttributeValue<string>("hil_name") : "";
                                    if (tempServiceName != "")
                                    {
                                        serviceName = serviceName + tempServiceName + ", ";
                                    }
                                }
                            }
                            if (serviceName != "")
                                serviceName = serviceName.Substring(0, serviceName.Length - 2);
                        }
                        transcationHistory.PlanDuration = servicelineEntCol.Entities.Count.ToString() + " Services";
                        transcationHistory.ProductName = serviceName;
                        lstTranscationHistories.Add(transcationHistory);
                    }
                }
                else
                {
                    return new List<TranscationHistoryWithOrder>();
                }
                return lstTranscationHistories;
            }
            catch (Exception)
            {
                return new List<TranscationHistoryWithOrder>();
            }
        }
        private static List<AMCProductWarranty> GetWarrantyDetails(Guid ProductGuid, ICrmService service)
        {
            List<AMCProductWarranty> lstWarrantyLineInfo = new List<AMCProductWarranty>();
            QueryExpression Query = new QueryExpression("hil_unitwarranty");
            Query.ColumnSet = new ColumnSet("hil_warrantystartdate", "hil_warrantyenddate", "hil_warrantytemplate");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            Query.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, ProductGuid);
            Query.Criteria.AddCondition("hil_warrantyenddate", ConditionOperator.GreaterEqual, DateTime.Now);
            EntityCollection entColl = service.RetrieveMultiple(Query);

            if (entColl.Entities.Count > 0)
            {
                foreach (Entity entwarranty in entColl.Entities)
                {
                    AMCProductWarranty objWarrantyLineInfo = new AMCProductWarranty();

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
                            objWarrantyLineInfo.WarrantyType = WarrantyType;
                        }
                    }
                    lstWarrantyLineInfo.Add(objWarrantyLineInfo);
                }
            }
            return lstWarrantyLineInfo;
        }
        public decimal GetDiscountValue(ICrmService service, Guid ModelId, Guid sourceId, int ProductAgeing, Guid ProductCategoryId, Guid ProductSubcategoryId, Guid Planid, Guid entState, Guid entSalesOffice)
        {
            decimal DiscPer = 0.00M;
            string _applicableOn = DateTime.Now.ToString("yyyy-MM-dd");
            string fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                                <entity name='hil_amcdiscountmatrix'>
                                <attribute name='hil_discper' />
                                <order attribute='hil_product' descending='true' />
                                <order attribute='hil_salesoffice' descending='true' />
                                <order attribute='hil_state' descending='true' />
                                <order attribute='hil_model' descending='true' />
                                <order attribute='hil_productcategory' descending='true' />
                                <order attribute='hil_productsubcategory' descending='true' />
                                <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_validfrom' operator='on-or-before' value='{_applicableOn}' />
                                <condition attribute='hil_validto' operator='on-or-after' value='{_applicableOn}' />
                                <condition attribute='hil_productaegingstart' operator='le' value='{ProductAgeing}' />
                                <condition attribute='hil_productageingend' operator='ge' value='{ProductAgeing}' />
                                <condition attribute='hil_appliedto' operator='eq' value='{sourceId}' />
                                <filter type='or'>
                                <condition attribute='hil_productcategory' operator='eq' value='{ProductCategoryId}' />
                                <condition attribute='hil_productcategory' operator='null' />
                                </filter>
                                <filter type='or'>
                                <condition attribute='hil_productsubcategory' operator='eq' value='{ProductSubcategoryId}' />
                                <condition attribute='hil_productsubcategory' operator='null' />
                                </filter>
                                <filter type='or'>
                                <condition attribute='hil_model' operator='eq' value='{ModelId}' />
                                <condition attribute='hil_model' operator='null' />
                                </filter>
                                <filter type='or'>
                                <condition attribute='hil_state' operator='null' />
                                <condition attribute='hil_state' operator='eq' value='{entState}' />
                                </filter>
                                <filter type='or'>
                                <condition attribute='hil_salesoffice' operator='null' />
                                <condition attribute='hil_salesoffice' operator='eq' value='{{90503976-8FD1-EA11-A813-000D3AF0563C}}' />
                                <condition attribute='hil_salesoffice' operator='eq' value='{entSalesOffice}' />
                                </filter>
                                <filter type='or'>
                                <condition attribute='hil_product' operator='eq' value='{Planid}' />
                                <condition attribute='hil_product' operator='null' />
                                </filter>
                                </filter>
                                </entity>
                                </fetch>";

            EntityCollection entamcdiscountmatrix = service.RetrieveMultiple(new FetchExpression(fetchQuery));
            if (entamcdiscountmatrix.Entities.Count > 0)
            {
                if (entamcdiscountmatrix.Entities[0].Contains("hil_discper"))
                {
                    DiscPer = entamcdiscountmatrix.Entities[0].GetAttributeValue<Decimal>("hil_discper");
                }
            }
            return Math.Round(DiscPer, 2);
        }
        public ResInvoiceInfo PayNow(Guid _orderId, Guid AMCPlanID, decimal Amount, ConsumerProfile _consumerProfile, ICrmService service)
        {
            SendPayNowRequest req = new SendPayNowRequest();
            RequestData requestData = new RequestData();
            ResInvoiceInfo _SendPayNowReponse = new ResInvoiceInfo();

            string _txnId = string.Empty;
            Guid _paymentReceiptId = Guid.Empty;

            (_paymentReceiptId, _txnId) = CreatePaymentReceipt(service, _orderId, _consumerProfile.emailId, _consumerProfile.mobileNumber, Amount, _consumerProfile.branchMemorandumCode);

            Entity ent = service.Retrieve("product", AMCPlanID, new ColumnSet("name"));
            string _productinfo = ent.Contains("name") ? ent.GetAttributeValue<string>("name") : "";
            requestData.amount = Math.Round(Amount, 2).ToString();

            requestData.txnid = _txnId;
            requestData.firstname = _consumerProfile.fullName;
            requestData.email = _consumerProfile.emailId;
            requestData.productinfo = _consumerProfile.branchMemorandumCode;
            requestData.udf2 = "D365";
            req.businessType = "B2C";
            req.paymentgateway_type = "PayUBiz";
            req.IM_PROJECT = "D365";
            req.RequestData = requestData;

            IntegrationConfiguration inconfig = CommonMethods.GetIntegrationConfiguration(_crmService, _PayNowKey);

            var data = new StringContent(JsonConvert.SerializeObject(req), Encoding.UTF8, "application/json");

            HttpClient client = new HttpClient();
            var byteArray = Encoding.ASCII.GetBytes(inconfig.userName + ":" + inconfig.password);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
            HttpResponseMessage response = client.PostAsync(inconfig.url, data).Result;
            if (response.IsSuccessStatusCode)
            {
                var obj = JsonConvert.DeserializeObject<PayNowResponse>(response.Content.ReadAsStringAsync().Result);
                if (obj.status == "Success")
                {
                    Entity _entPaymentReceipt = new Entity("hil_paymentreceipt", _paymentReceiptId);
                    _entPaymentReceipt["hil_paymentstatus"] = new OptionSetValue(1);
                    _entPaymentReceipt["hil_response"] = "Success";
                    _entPaymentReceipt["hil_tokenexpireson"] = DateTime.Now.AddMinutes(340);
                    service.Update(_entPaymentReceipt);

                    _SendPayNowReponse.TransactionID = _txnId;
                    _SendPayNowReponse.productinfo = _consumerProfile.branchMemorandumCode;
                    _SendPayNowReponse.HashCode = obj.HashCode;
                    _SendPayNowReponse.salt = obj.salt;
                    _SendPayNowReponse.key = obj.key;
                    _SendPayNowReponse.InvoiceID = _orderId;
                    _SendPayNowReponse.ResultStatus = true;
                    _SendPayNowReponse.ResultMessage = obj.message;
                }
                else
                {
                    _SendPayNowReponse.ResultStatus = false;
                    _SendPayNowReponse.ResultMessage = obj.message;
                }
            }
            return _SendPayNowReponse;
        }
        private (Guid, string) CreatePaymentReceipt(ICrmService service, Guid _orderId, string _email, string _mobileNumber, decimal _amount, string _mamorandumCode)
        {
            Guid _paymentReceiptId = Guid.Empty;
            string _transactionId = string.Empty;
            Entity paymentReceipt = new Entity("hil_paymentreceipt");
            paymentReceipt["hil_orderid"] = new EntityReference("salesorder", _orderId);
            paymentReceipt["hil_email"] = _email;
            paymentReceipt["hil_mobilenumber"] = _mobileNumber;
            paymentReceipt["hil_amount"] = new Money(_amount);
            paymentReceipt["hil_memorandumcode"] = _mamorandumCode;
            try
            {
                _paymentReceiptId = service.Create(paymentReceipt);
                _transactionId = service.Retrieve("hil_paymentreceipt", _paymentReceiptId, new ColumnSet("hil_transactionid")).Attributes["hil_transactionid"].ToString();
            }
            catch (Exception ex)
            {
                throw new Exception($"ERROR! {ex.Message}");
            }
            return (_paymentReceiptId, _transactionId);
        }
        #endregion
    }
}
