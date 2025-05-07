using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.AMC
{
    public class InitiateAMCPayment : IPlugin
    {
        private static Guid PriceLevelForFGsale = new Guid("c05e7837-8fc4-ed11-b597-6045bdac5e78");//AMC Ominichannnel
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            string[] Source = { "6" };

            string LoginUserId = Convert.ToString(context.InputParameters["LoginUserId"]);
            string UserToken = Convert.ToString(context.InputParameters["UserToken"]);

            string jsonString = Convert.ToString(context.InputParameters["reqdata"]);
            var data = JsonSerializer.Deserialize<InitiatePaymentParam>(jsonString);
            string AMCPlanID = data.AMCPlanID;
            string AssestId = data.AssestId;
            string MobileNumber = LoginUserId;
            string AddressID = data.AddressID;
            string Price = data.Price.ToString();
            string DateofPurchase = data.DOP;
            string SourceType = data.SourceType;
            DateTime DOP;
            if (!APValidate.IsvalidGuid(AMCPlanID))
            {
                string msg = string.IsNullOrWhiteSpace(AMCPlanID) ? "AMC Plan id is required." : "Invalid AMC Plan id.";
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = responnse;
                return;
            }
            if (!APValidate.IsvalidGuid(AssestId))
            {
                string msg = string.IsNullOrWhiteSpace(AssestId) ? "Customer Assest id is required" : "Invalid Customer Assit id";
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = responnse;
                return;
            }
            if (!APValidate.IsValidMobileNumber(MobileNumber))
            {

                string msg = string.IsNullOrWhiteSpace(MobileNumber) ? "Mobile number is required." : "Invalid mobile number.";
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = responnse;
                return;
            }

            if (!APValidate.IsvalidGuid(AddressID))
            {
                string msg = string.IsNullOrWhiteSpace(AddressID) ? "Address guid required" : "Invalid address guid";
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = responnse;
                return;
            }
            if (!APValidate.IsDecimal(Price))
            {
                string msg = string.IsNullOrWhiteSpace(Price) ? "Price value is required" : "Invalid Price value.It's should be greater than 0";
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = responnse;
                return;
            }

            if (!DateTime.TryParseExact(DateofPurchase, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DOP))
            {
                string msg = string.IsNullOrWhiteSpace(DateofPurchase) ? "Date of purchase is required" : "Invalid date format. It should be (yyyy-MM-dd)";
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = responnse;
                return;
            }
            if (!Source.Contains(SourceType))
            {
                string msg = string.IsNullOrWhiteSpace(SourceType) ? "Source type is required." : "Invalid source type.";
                var responnse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = responnse;
                return;
            }
            InitiatePaymentParam objparam = new InitiatePaymentParam();
            objparam.AMCPlanID = AMCPlanID;
            objparam.AssestId = AssestId;
            objparam.AddressID = AddressID;
            objparam.DOP = DOP.Date.ToString("yyyy-MM-dd");
            objparam.Price = Price;
            objparam.SourceType = SourceType;

            var response = InitiatePayment(service, objparam, MobileNumber);
            dynamic result;

            if (response.Item2.StatusCode == (int)HttpStatusCode.OK)
                result = JsonSerializer.Serialize(response.Item1);
            else
                result = JsonSerializer.Serialize(response.Item2);
            context.OutputParameters["data"] = result;
            return;
        }
        public (InitiatePaymentRes, RequestStatus) InitiatePayment(IOrganizationService _crmService, InitiatePaymentParam IPParam, string MobileNumber)
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
                string _requestData = JsonSerializer.Serialize(IPParam);
                if (_crmService != null)
                {
                    if (!APValidate.IsValidMobileNumber(MobileNumber))
                    {
                        return (initiatePaymentRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = "Invalid Mobile Number"
                        });
                    }
                    if (IPParam.SourceType != "6")
                    {
                        string msg = "Source type is required.";
                        if (!string.IsNullOrWhiteSpace(IPParam.SourceType))
                            msg = "Invalid Source Type";
                        return (initiatePaymentRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = msg
                        });
                    }
                    if (!APValidate.IsvalidGuid(IPParam.AMCPlanID))
                    {
                        string msg = "AMC Plan id required.";
                        if (!string.IsNullOrWhiteSpace(IPParam.AMCPlanID))
                            msg = "Invalid AMC Plan Guid.";
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
                                Message = "Invalid AMC Plan Guid."
                            });
                        }
                        else
                            _entRefAMCPlan = _entObj.ToEntityReference();
                    }
                    if (!APValidate.IsvalidGuid(IPParam.AssestId))
                    {
                        string msg = "Assest Id is required";
                        if (!string.IsNullOrWhiteSpace(IPParam.AssestId))
                            msg = "Invalid Customer Assit Id";
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
                                Message = "Invalid Customer Assit Id"
                            });
                        }
                    }
                    if (!APValidate.IsvalidGuid(IPParam.AddressID))
                    {
                        string msg = "Address Guid required";
                        if (!string.IsNullOrWhiteSpace(IPParam.AddressID))
                            msg = "Invalid Address Guid";
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
                                Message = "Invalid Address Guid"
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
                    if (Convert.ToDecimal(IPParam.Price) < 1)
                    {
                        return (initiatePaymentRes, new RequestStatus()
                        {
                            StatusCode = (int)HttpStatusCode.BadRequest,
                            Message = "Price is required and should be greater than 0"
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
                    Guid CustomerGuid = getCustomerGuid(_crmService, MobileNumber);
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
                    EntityReference _entRefServiceAccount = new EntityReference("account", new Guid("d166ba69-65da-ec11-a7b5-6045bdad2a19"));
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
                                _paymentStatus = GetPaymentStatus(_crmService, _transactionId, _entRefPaymentTransId.Id, "Send Payment Link");
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
                        ResInvoiceInfo resInvoiceInfo = PayNow(_entRefOrder.Id, _entRefAMCPlan.Id, Convert.ToDecimal(IPParam.Price), _consumerProfile, _crmService);
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
                        initiatePaymentRes.Surl = "https://havellscrmwebapp-prod.azurewebsites.net/Web/PaymentSuccess.html";
                        initiatePaymentRes.Furl = "https://havellscrmwebapp-prod.azurewebsites.net/Web/PaymentFail.html";

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
                        Message = "D365 service unavailable."
                    });
                }
            }
            catch (Exception ex)
            {
                return (initiatePaymentRes, new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "D365 internal server error : " + ex.Message.ToUpper()
                });
            }
        }
        private static Guid getCustomerGuid(IOrganizationService _service, string MobileNumber)
        {
            QueryExpression queryExp = new QueryExpression("contact");
            queryExp.ColumnSet = new ColumnSet(false);
            ConditionExpression condExp = new ConditionExpression("mobilephone", ConditionOperator.Equal, MobileNumber);
            queryExp.Criteria.AddCondition(condExp);
            condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0);//Active
            queryExp.Criteria.AddCondition(condExp);
            EntityCollection entCol = _service.RetrieveMultiple(queryExp);
            if (entCol.Entities.Count > 0)
            {
                return entCol.Entities[0].Id;
            }
            else
            {
                return Guid.Empty;
            }
        }
        private static IntegrationConfiguration GetIntegrationConfiguration(IOrganizationService _service, string name)
        {
            IntegrationConfiguration inconfig = new IntegrationConfiguration();
            QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
            qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
            qsCType.NoLock = true;
            qsCType.Criteria = new FilterExpression(LogicalOperator.And);
            qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, name);
            Entity integrationConfiguration = _service.RetrieveMultiple(qsCType)[0];
            inconfig.url = integrationConfiguration.GetAttributeValue<string>("hil_url");
            inconfig.userName = integrationConfiguration.GetAttributeValue<string>("hil_username");
            inconfig.password = integrationConfiguration.GetAttributeValue<string>("hil_password");
            return inconfig;
        }
        private static OptionSetValue GetPaymentStatus(IOrganizationService service, string _transactionId, Guid _paymentReceiptId, string _SendPaymentLink)
        {
            OptionSetValue _paymentStatus = null;
            StatusRequest reqParm = new StatusRequest();
            reqParm.PROJECT = "D365";
            reqParm.command = "verify_payment";
            reqParm.var1 = _transactionId;

            IntegrationConfiguration inconfig = GetIntegrationConfiguration(service, _SendPaymentLink);
            var data = new StringContent(JsonSerializer.Serialize(reqParm), Encoding.UTF8, "application/json");
            HttpClient client = new HttpClient();
            var byteArray = Encoding.ASCII.GetBytes(inconfig.userName + ":" + inconfig.password);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
            HttpResponseMessage response = client.PostAsync(inconfig.url, data).Result;
            if (response.IsSuccessStatusCode)
            {
                var obj = JsonSerializer.Deserialize<StatusResponse>(response.Content.ReadAsStringAsync().Result);

                foreach (var item in obj.transaction_details)
                {
                    Entity Paymentreceipt = new Entity("hil_paymentreceipt", _paymentReceiptId);
                    if (obj.transaction_details[0].mode != null)
                    {
                        Paymentreceipt["hil_paymentmode"] = obj.transaction_details[0].mode.ToString();
                    }
                    if (obj.transaction_details[0].bank_ref_num != null)
                    {
                        Paymentreceipt["hil_bankreferenceid"] = obj.transaction_details[0].bank_ref_num.ToString();
                    }
                    if (obj.transaction_details[0].addedon != null)
                    {
                        Paymentreceipt["hil_receiptdate"] = DateTime.Parse(obj.transaction_details[0].addedon);
                    }
                    if (obj.transaction_details[0].error_Message != null)
                    {
                        Paymentreceipt["hil_response"] = obj.transaction_details[0].error_Message.ToString();
                    }
                    string status = obj.transaction_details[0].status;
                    if (status == "not initiated")
                    {
                        Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(1);
                        _paymentStatus = new OptionSetValue(1);
                    }
                    else if (status == "success")
                    {
                        Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(4);
                        _paymentStatus = new OptionSetValue(4);
                    }
                    else if (status == "pending")
                    {
                        Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(3);
                        _paymentStatus = new OptionSetValue(3);
                    }
                    else if (status == "failure")
                    {
                        Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(2);
                        _paymentStatus = new OptionSetValue(2);
                    }
                    else
                    {
                        _paymentStatus = null;
                    }
                    service.Update(Paymentreceipt);
                }
            }
            return _paymentStatus;
        }
        private ResInvoiceInfo PayNow(Guid _orderId, Guid AMCPlanID, decimal Amount, ConsumerProfile _consumerProfile, IOrganizationService service)
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

            IntegrationConfiguration inconfig = GetIntegrationConfiguration(service, "Pay Now");

            var data = new StringContent(JsonSerializer.Serialize(req), Encoding.UTF8, "application/json");

            HttpClient client = new HttpClient();
            var byteArray = Encoding.ASCII.GetBytes(inconfig.userName + ":" + inconfig.password);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
            HttpResponseMessage response = client.PostAsync(inconfig.url, data).Result;
            if (response.IsSuccessStatusCode)
            {
                var obj = JsonSerializer.Deserialize<PayNowResponse>(response.Content.ReadAsStringAsync().Result);
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
        private (Guid, string) CreatePaymentReceipt(IOrganizationService service, Guid _orderId, string _email, string _mobileNumber, decimal _amount, string _mamorandumCode)
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

        #region Initiate Payment
        public class InitiatePaymentParam
        {
            public string AMCPlanID { get; set; }
            public string AssestId { get; set; }
            public string DOP { get; set; }
            public string Price { get; set; }
            public string AddressID { get; set; }
            public string SourceType { get; set; }
        }
        public class InitiatePaymentRes : TokenExpires
        {
            public string TransactionID { get; set; }
            public string InvoiceID { get; set; }
            public string HashName { get; set; }
            public string Key { get; set; }
            public string Salt { get; set; }
            public string MobileNumber { get; set; }
            public string MamorandumCode { get; set; }
            public string Emailid { get; set; }
            public string UserName { get; set; }
            public string Surl { get; set; }
            public string Furl { get; set; }
        }
        public class SendPayNowRequest
        {
            public string businessType { get; set; }
            public string paymentgateway_type { get; set; }
            public string IM_PROJECT { get; set; }
            public RequestData RequestData { get; set; }
        }
        public class RequestData
        {
            public string txnid { get; set; }
            public string amount { get; set; }
            public string productinfo { get; set; }
            public string firstname { get; set; }
            public string email { get; set; }
            public string udf1 { get; set; }
            public string udf2 { get; set; }
            public string udf3 { get; set; }
            public string udf4 { get; set; }
            public string udf5 { get; set; }

        }
        public class PayNowResponse
        {
            public string HashCode { get; set; }
            public string key { get; set; }
            public string salt { get; set; }
            public string status { get; set; }
            public string message { get; set; }

        }
        public class ConsumerProfile
        {
            public string mobileNumber { get; set; }
            public string fullName { get; set; }
            public string emailId { get; set; }
            public string state { get; set; }
            public string pinCode { get; set; }
            public string salesOffice { get; set; }
            public string branchMemorandumCode { get; set; }
            public string fullAddress { get; set; }
        }
        public class ResInvoiceInfo
        {
            public string TransactionID { get; set; }
            public Guid InvoiceID { get; set; }
            public string productinfo { get; set; }
            public string HashCode { get; set; }
            public string key { get; set; }
            public string salt { get; set; }
            public bool ResultStatus { get; set; }
            public string ResultMessage { get; set; }

        }
        public class IntegrationConfiguration
        {
            public string url { get; set; }
            public string userName { get; set; }
            public string password { get; set; }
        }
        public class StatusRequest
        {
            public string PROJECT { get; set; }
            public string command { get; set; }
            public string var1 { get; set; }
        }
        public class StatusResponse
        {
            public int status { get; set; }
            public string msg { get; set; }
            public List<TransactionDetail> transaction_details { get; set; }
        }
        public class TransactionDetail
        {
            public object bank_ref_num { get; set; }
            public string addedon { get; set; }
            public string error_Message { get; set; }
            public string mode { get; set; }
            public string status { get; set; }
        }
        public class TokenExpires
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
        }
        public class RequestStatus
        {
            public int StatusCode { get; set; }
            public string Message { get; set; }
        }

        #endregion
    }
}
