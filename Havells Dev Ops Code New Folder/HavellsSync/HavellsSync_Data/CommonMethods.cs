using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Text;

namespace HavellsSync_Data
{
    public static class CommonMethods
    {
        public static Entity GetServiceDetails(ICrmService _service, Guid ProductId)
        {
            Entity entProductcatalog = new Entity();
            string xmlqery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' top='1' distinct='false'>
                  <entity name='hil_productcatalog'>
                    <attribute name='hil_productcatalogid' />
                    <attribute name='hil_name' />
                    <attribute name='hil_plantclink' />
                    <attribute name='hil_productcode' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                      <condition attribute='hil_productcode' operator='eq' value='{ProductId.ToString()}' />
                    </filter>
                  </entity>
                </fetch>";
            EntityCollection entitycol = _service.RetrieveMultiple(new FetchExpression(xmlqery));
            if (entitycol.Entities.Count > 0)
            {
                entProductcatalog = entitycol.Entities[0];
            }
            return entProductcatalog;
        }
        public static DateTime? getExpiryTime(ICrmService _service, string SessionId)
        {
            QueryExpression queryExp = new QueryExpression("hil_consumerloginsession");
            queryExp.ColumnSet = new ColumnSet("hil_expiredon", "hil_consumerloginsessionid", "hil_origin");
            ConditionExpression condExp = new ConditionExpression("hil_consumerloginsessionid", ConditionOperator.Equal, SessionId);
            queryExp.Criteria.AddCondition(condExp);
            condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0);
            queryExp.Criteria.AddCondition(condExp);
            EntityCollection entCol = _service.RetrieveMultiple(queryExp);
            if (entCol.Entities.Count > 0)
            {
                return entCol.Entities[0].GetAttributeValue<DateTime>("hil_expiredon").AddMinutes(330);
            }
            else
            {
                return null;
            }
        }
        public static string getSourceType(ICrmService _service, string SessionId)
        {
            Entity Consumerloginsession = _service.Retrieve("hil_consumerloginsession", new Guid(SessionId), new ColumnSet("hil_origin"));
            return Consumerloginsession.Contains("hil_origin") ? Consumerloginsession.GetAttributeValue<string>("hil_origin") : "";
        }
        public static bool IsvalidGuid(ICrmService service, string entity, Guid guid)
        {
            bool result = false;
            Entity ent = service.Retrieve(entity, guid, new ColumnSet(false));
            if (ent != null)
            {
                result = true;
            }
            return result;
        }
        public static IntegrationConfiguration GetIntegrationConfiguration(ICrmService _service, string name)
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
        public static bool CheckIfExistingSerialNumber(ICrmService service, string Serial)
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
        public static Entity CheckIfExistingSerialNumberWithDetails(ICrmService service, string Serial)
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
        public static void AttachNotes(ICrmService service, byte[] NoteByte, Guid Asset, int fileType, string entityType)
        {
            try
            {
                if (NoteByte != null && Asset != Guid.Empty)//int can't be null
                {
                    Entity An = new Entity("annotation");
                    An["documentbody"] = Convert.ToBase64String(NoteByte);
                    if (fileType == 0)
                    {
                        An["mimetype"] = "image/jpeg";
                        An["filename"] = "invoice.jpeg";
                    }
                    else if (fileType == 1)
                    {
                        An["mimetype"] = "image/png";
                        An["filename"] = "invoice.png";
                    }
                    if (fileType == 2)
                    {
                        An["mimetype"] = "application/pdf";
                        An["filename"] = "invoice.pdf";
                    }
                    else if (fileType == 3)
                    {
                        An["mimetype"] = "application/doc";
                        An["filename"] = "invoice.doc";
                    }
                    else if (fileType == 4)
                    {
                        An["mimetype"] = "application/tiff";
                        An["filename"] = "invoice.tiff";
                    }
                    else if (fileType == 5)
                    {
                        An["mimetype"] = "application/gif";
                        An["filename"] = "invoice.gif";
                    }
                    else if (fileType == 6)
                    {
                        An["mimetype"] = "application/bmp";
                        An["filename"] = "invoice.bmp";
                    }
                    An["objectid"] = new EntityReference(entityType, Asset);
                    An["objecttypecode"] = entityType;
                    service.Create(An);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(CommonMessage.ErrorattachingnotesMsg + ex.Message);
            }
        }
        public static Guid GetGuidbyName98(String sEntityName, String sFieldName, String sFieldValue, ICrmService service, int iStatusCode = 0)
        {
            Guid fsResult = Guid.Empty;
            try
            {
                QueryExpression qe = new QueryExpression(sEntityName);
                qe.Criteria.AddCondition(sFieldName, ConditionOperator.Equal, sFieldValue);
                qe.AddOrder("createdon", OrderType.Descending);
                if (iStatusCode >= 0)
                {
                    qe.Criteria.AddCondition("statecode", ConditionOperator.Equal, iStatusCode);
                }
                EntityCollection enColl = service.RetrieveMultiple(qe);
                if (enColl.Entities.Count > 0)
                {
                    fsResult = enColl.Entities[0].Id;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(CommonMessage.Havells_PluginMsg + ex.Message);
            }
            return fsResult;
        }
        public static Guid getCustomerGuid(ICrmService _service, string MobileNumber)
        {
            //Entity Consumerloginsession = _service.Retrieve("hil_consumerloginsession", new Guid(LoginUserId), new ColumnSet("hil_name"));
            //string UserCode = Consumerloginsession.Contains("hil_name") ? Consumerloginsession.GetAttributeValue<string>("hil_name") : "";
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
        public static string getTransactionStatus_Old(ICrmService _service, string TxnId, Guid InvoiceId, string _SendPaymentLink)
        {
            string Status = "failure";
            StatusRequest reqParm = new StatusRequest();
            reqParm.PROJECT = "D365";
            reqParm.command = "verify_payment";
            reqParm.var1 = TxnId;
            try
            {
                QueryExpression query = new QueryExpression("hil_paymentstatus");
                query.ColumnSet = new ColumnSet("hil_paymentstatusid");
                query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, TxnId);
                EntityCollection entpaymentstatus = _service.RetrieveMultiple(query);

                IntegrationConfiguration inconfig = CommonMethods.GetIntegrationConfiguration(_service, _SendPaymentLink);
                var data = new StringContent(JsonConvert.SerializeObject(reqParm), Encoding.UTF8, "application/json");
                HttpClient client = new HttpClient();
                var byteArray = Encoding.ASCII.GetBytes(inconfig.userName + ":" + inconfig.password);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                HttpResponseMessage response = client.PostAsync(inconfig.url, data).Result;
                if (response.IsSuccessStatusCode)
                {
                    var obj = JsonConvert.DeserializeObject<StatusResponse>(response.Content.ReadAsStringAsync().Result);
                    foreach (var item in obj.transaction_details)
                    {
                        Entity statusPayment = new Entity("hil_paymentstatus");
                        statusPayment["hil_mihpayid"] = item.mihpayid;
                        statusPayment["hil_request_id"] = item.request_id;
                        statusPayment["hil_bank_ref_num"] = item.bank_ref_num;
                        statusPayment["hil_amt"] = item.amt;
                        statusPayment["hil_transaction_amount"] = item.transaction_amount;
                        statusPayment["hil_name"] = item.txnid == null ? TxnId : item.txnid;
                        statusPayment["hil_additional_charges"] = item.additional_charges;
                        statusPayment["hil_productinfo"] = item.productinfo;
                        statusPayment["hil_firstname"] = item.firstname;
                        statusPayment["hil_bankcode"] = item.bankcode;
                        statusPayment["hil_udf1"] = item.udf1;
                        statusPayment["hil_udf2"] = item.udf2;
                        statusPayment["hil_udf3"] = item.udf3;
                        statusPayment["hil_udf4"] = item.udf4;
                        statusPayment["hil_udf5"] = item.udf5;
                        statusPayment["hil_field2"] = item.field2;
                        statusPayment["hil_field9"] = item.field9;
                        statusPayment["hil_error_code"] = item.error_code;
                        statusPayment["hil_addedon"] = item.addedon;
                        statusPayment["hil_payment_source"] = item.payment_source;
                        statusPayment["hil_card_type"] = item.card_type;
                        statusPayment["hil_error_message"] = item.error_Message;
                        statusPayment["hil_net_amount_debit"] = item.net_amount_debit;
                        statusPayment["hil_disc"] = item.disc;
                        statusPayment["hil_mode"] = item.mode;
                        statusPayment["hil_pg_type"] = item.PG_TYPE;
                        statusPayment["hil_card_no"] = item.card_no;
                        statusPayment["hil_paymentstatus"] = item.status;
                        statusPayment["hil_unmappedstatus"] = item.unmappedstatus;
                        statusPayment["hil_merchant_utr"] = item.Merchant_UTR;
                        statusPayment["hil_settled_at"] = item.Settled_At;
                        if (entpaymentstatus.Entities.Count > 0)
                        {
                            statusPayment.Id = entpaymentstatus[0].Id;
                            _service.Update(statusPayment);
                        }
                        else
                        {
                            _service.Create(statusPayment);
                        }
                        Status = item.status.ToLower();
                    }
                    if (Status.ToLower() == "success".ToLower())
                    {
                        SetStateRequest req = new SetStateRequest();
                        req.State = new OptionSetValue(2);
                        req.Status = new OptionSetValue(100001);
                        req.EntityMoniker = new EntityReference("invoice", InvoiceId);
                        _service.Execute(req);

                        QueryExpression Query = new QueryExpression("msdyn_paymentdetail");
                        Query.ColumnSet = new ColumnSet(false);
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, TxnId);
                        EntityCollection FoundPaymentDetails = _service.RetrieveMultiple(Query);

                        req = new SetStateRequest();
                        req.State = new OptionSetValue(0);
                        req.Status = new OptionSetValue(910590000);
                        req.EntityMoniker = FoundPaymentDetails[0].ToEntityReference();
                        _service.Execute(req);
                    }
                    else
                    {
                        if (Status.ToLower() == "failure".ToLower())
                        {
                            QueryExpression Query = new QueryExpression("msdyn_paymentdetail");
                            Query.ColumnSet = new ColumnSet(false);
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, TxnId);
                            EntityCollection FoundPaymentDetails = _service.RetrieveMultiple(Query);

                            Entity entity = new Entity(FoundPaymentDetails[0].LogicalName, FoundPaymentDetails[0].Id);
                            entity["msdyn_paymentamount"] = new Money(0);
                            _service.Update(entity);

                            SetStateRequest req = new SetStateRequest();
                            req.State = new OptionSetValue(0);
                            req.Status = new OptionSetValue(910590001);
                            req.EntityMoniker = FoundPaymentDetails[0].ToEntityReference();
                            _service.Execute(req);
                        }
                    }
                }
                return Status;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string getTransactionStatus(ICrmService _service, Guid entPaymentReceiptId, string TxnId, Guid InvoiceId, string _SendPaymentLink)
        {
            string Status = "Failed";
            StatusRequest reqParm = new StatusRequest();
            reqParm.PROJECT = "D365";
            reqParm.command = "verify_payment";
            reqParm.var1 = TxnId;
            try
            {
                IntegrationConfiguration inconfig = CommonMethods.GetIntegrationConfiguration(_service, _SendPaymentLink);
                var data = new StringContent(JsonConvert.SerializeObject(reqParm), Encoding.UTF8, "application/json");
                HttpClient client = new HttpClient();
                var byteArray = Encoding.ASCII.GetBytes(inconfig.userName + ":" + inconfig.password);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                HttpResponseMessage response = client.PostAsync(inconfig.url, data).Result;
                if (response.IsSuccessStatusCode)
                {
                    var obj = JsonConvert.DeserializeObject<StatusResponse>(response.Content.ReadAsStringAsync().Result);
                    foreach (var item in obj.transaction_details)
                    {
                        Entity Paymentreceipt = new Entity("hil_paymentreceipt", entPaymentReceiptId);
                        if (!string.IsNullOrWhiteSpace(item.bank_ref_num))
                        {
                            Paymentreceipt["hil_bankreferenceid"] = item.bank_ref_num;
                        }
                        if (!string.IsNullOrWhiteSpace(item.addedon))
                        {
                            Paymentreceipt["hil_receiptdate"] = DateTime.Parse(item.addedon);
                        }
                        if (!string.IsNullOrWhiteSpace(item.amt))
                        {
                            Paymentreceipt["hil_amount"] = Decimal.Parse(item.amt);
                        }
                        if (!string.IsNullOrWhiteSpace(item.error_Message))
                        {
                            Paymentreceipt["hil_response"] = item.error_Message;
                        }
                        if (item.status.ToLower() == "not initiated" || item.status.ToLower() == "not found")
                        {
                            Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(1);
                            Status = "Pending";
                        }
                        else if (item.status.ToLower() == "pending")
                        {
                            Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(3);
                            Status = "Pending";
                        }
                        else if (item.status.ToLower() == "success")
                        {
                            Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(4);
                            Status = "Success";
                        }
                        else //failure
                        {
                            Paymentreceipt["hil_paymentstatus"] = new OptionSetValue(2);
                        }
                        _service.Update(Paymentreceipt);
                    }
                }
                return Status;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        public static string getCustomerMobile(ICrmService _service, string SessionId)
        {
            Entity Consumerloginsession = _service.Retrieve("hil_consumerloginsession", new Guid(SessionId), new ColumnSet("hil_name"));
            return Consumerloginsession.Contains("hil_name") ? Consumerloginsession.GetAttributeValue<string>("hil_name") : "";
        }
        public static string getSerialNumberByAssestId(ICrmService _service, Guid AssetId)
        {
            Entity customerasset = _service.Retrieve("msdyn_customerasset", AssetId, new ColumnSet("msdyn_name"));
            return customerasset.Contains("msdyn_name") ? customerasset.GetAttributeValue<string>("msdyn_name") : "";
        }
        public static bool IsValidModelNumber(ICrmService _service, string ModelNumber)
        {
            bool IsValid = false;
            string fetchXmlQuery = $@"<fetch top='1' version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
               <entity name='product'>
                 <attribute name='name' />
                 <attribute name='hil_materialgroup' />
                 <attribute name='hil_division' />
                 <order attribute='productnumber' descending='false' />
                 <filter type='and'>
                   <condition attribute='name' operator='eq' value='{ModelNumber}' />
                 </filter>
               </entity>
             </fetch>";
            EntityCollection entityColProduct = _service.RetrieveMultiple(new FetchExpression(fetchXmlQuery));
            if (entityColProduct.Entities.Count > 0)
            {
                IsValid = true;
            }
            return IsValid;
        }
        public static string ValidateTwoDates(DateTime ToDate, DateTime FromDate)
        {
            if (ToDate >= FromDate)
            {
                return "";
            }
            return "To Date should be greater than or equal to From Date.";
        }
        public static OptionSetValue GetPaymentStatus(ICrmService service, string _transactionId, Guid _paymentReceiptId, string _SendPaymentLink)
        {
            OptionSetValue _paymentStatus = null;
            StatusRequest reqParm = new StatusRequest();
            reqParm.PROJECT = "D365";
            reqParm.command = "verify_payment";
            reqParm.var1 = _transactionId;

            IntegrationConfiguration inconfig = CommonMethods.GetIntegrationConfiguration(service, _SendPaymentLink);
            var data = new StringContent(JsonConvert.SerializeObject(reqParm), Encoding.UTF8, "application/json");
            HttpClient client = new HttpClient();
            var byteArray = Encoding.ASCII.GetBytes(inconfig.userName + ":" + inconfig.password);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
            HttpResponseMessage response = client.PostAsync(inconfig.url, data).Result;
            if (response.IsSuccessStatusCode)
            {
                var obj = JsonConvert.DeserializeObject<StatusResponse>(response.Content.ReadAsStringAsync().Result);

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
    }
}
