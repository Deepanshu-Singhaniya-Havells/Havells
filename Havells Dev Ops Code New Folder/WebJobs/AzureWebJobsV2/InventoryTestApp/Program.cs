using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IdentityModel.Metadata;
using System.Linq;
using System.Runtime.Serialization;
using System.Security.Cryptography;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

namespace InventoryTestApp
{
    internal class Program
    {
        static IOrganizationService _service;
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        static void Main(string[] args)
        {
            SMLY sMLY = new SMLY();
            sMLY.RunFunction(args);

            //_service = ConnectToCRM(string.Format(connStr, "https://havellscrmdev1.crm8.dynamics.com"));

            //Post_PurchaseOrderRecipt(_service, new EntityReference("msdyn_purchaseorderreceipt", new Guid("58d06523-c161-ee11-8df0-6045bdac5292")));
            //SendURLD365Request reqParm  = new SendURLD365Request();
            //reqParm.mobile = "8005084995";
            //reqParm.jobId = "e9dc4228-a45b-ee11-8df0-6045bdac526a";
            //reqParm.Amount = "1";
            //SendSMS(reqParm, _service);

            //PurchaseOrderModel purchaseOrderTypeModel = new PurchaseOrderModel();
            //purchaseOrderTypeModel.msdyn_vendor = new EntityReference("account", new Guid("1d401d27-310b-e911-a94e-000d3af06091"));
            //purchaseOrderTypeModel.hil_potype = new EntityReference("hil_inventorytransactiontype", PurchaseOrderType.Emergency);
            //purchaseOrderTypeModel.msdyn_requestedbyresource = new EntityReference("bookableresource", new Guid("8b1f7f5f-ab53-ee11-be6f-6045bdaa91c3"));
            //purchaseOrderTypeModel.msdyn_receivetowarehouse = new EntityReference("msdyn_warehouse", new Guid("3d7a6dde-a753-ee11-be6f-6045bdaa91c3"));
            //purchaseOrderTypeModel.msdyn_purchaseorderdate = Convert.ToDateTime("23/09/2023");
            //purchaseOrderTypeModel.msdyn_orderedby = new EntityReference("systemuser", new Guid("0d8887f1-0f84-ec11-8d21-6045bdaad3de"));
            //purchaseOrderTypeModel.ownerid = new EntityReference("systemuser", new Guid("0d8887f1-0f84-ec11-8d21-6045bdaad3de"));
            //purchaseOrderTypeModel.msdyn_workorder = new EntityReference("msdyn_workorder", new Guid("7b867834-a94c-ee11-be6f-6045bdaa91c3"));
            //Guid POID = createPO(_service, purchaseOrderTypeModel);
        }

        public static void Post_PurchaseOrderRecipt(IOrganizationService service, EntityReference _poReciptEntityRef)
        {
            try
            {
                string fetch = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                  <entity name=""msdyn_purchaseorderreceiptproduct"">
                                    <attribute name=""hil_billedquantity"" />
                                    <attribute name=""msdyn_name"" />
                                    <attribute name=""hil_sortquantity"" />
                                    <attribute name=""hil_freshquantity"" />
                                    <attribute name=""hil_damagequantity"" />
                                    <attribute name=""msdyn_purchaseorder"" />
                                    <attribute name=""ownerid"" />
                                    <attribute name=""msdyn_purchaseorderreceipt"" />
                                    <attribute name=""msdyn_purchaseorderproduct"" />
                                    <attribute name=""msdyn_associatetoworkorder"" />
                                    <attribute name=""msdyn_associatetowarehouse"" />
                                    <attribute name=""msdyn_purchaseorderreceiptproductid"" />
                                    <filter type=""and"">
                                      <condition attribute=""statuscode"" operator=""eq"" value=""1"" />
                                      <condition attribute=""hil_billedquantity"" operator=""not-null"" />
                                      <condition attribute=""msdyn_purchaseorderreceipt"" operator=""eq"" value=""{_poReciptEntityRef.Id}"" />
                                    </filter>
                                  </entity>
                                </fetch>";
                EntityCollection collection = service.RetrieveMultiple(new FetchExpression(fetch));
                int sortQty = 0, freshQty = 0, totalQty = 0, damageQty = 0;
                foreach (Entity entity in collection.Entities)
                {
                    sortQty = 0; freshQty = 0; totalQty = 0; damageQty = 0;
                    freshQty = entity.GetAttributeValue<int>("hil_freshquantity");
                    sortQty = entity.GetAttributeValue<int>("hil_sortquantity");
                    totalQty = entity.GetAttributeValue<int>("hil_billedquantity");
                    damageQty = entity.GetAttributeValue<int>("hil_damagequantity");
                    if ((damageQty + sortQty + freshQty) != totalQty)
                    {
                        throw new Exception("Quantity is mismached for Purchase Order Recipt Product " + entity["msdyn_name"]);
                    }
                }
                foreach (Entity entity in collection.Entities)
                {
                    try
                    {
                        if (entity.Contains("hil_freshquantity"))
                        {
                            Entity entity1 = new Entity(entity.LogicalName, entity.Id);
                            entity1["msdyn_quantity"] = (double)entity.GetAttributeValue<int>("hil_freshquantity");
                            service.Update(entity1);
                        }
                        if (entity.Contains("hil_damagequantity"))
                        {
                            Entity entity1 = new Entity(entity.LogicalName);
                            entity1["msdyn_purchaseorder"] = entity["msdyn_purchaseorder"];
                            entity1["ownerid"] = entity["ownerid"];
                            entity1["msdyn_purchaseorderreceipt"] = entity["msdyn_purchaseorderreceipt"];
                            entity1["msdyn_purchaseorderproduct"] = entity["msdyn_purchaseorderproduct"];
                            entity1["msdyn_associatetoworkorder"] = entity["msdyn_associatetoworkorder"];
                            entity1["msdyn_associatetowarehouse"] = RetriveBookableResource(service, entity.GetAttributeValue<EntityReference>("ownerid")).DefectiveWarehouse;
                            entity1["msdyn_quantity"] = (double)entity.GetAttributeValue<int>("hil_damagequantity");
                            service.Create(entity1);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static RetriveBookableResourceResponse RetriveBookableResource(IOrganizationService service, EntityReference _userID)
        {
            RetriveBookableResourceResponse bookableResource = new RetriveBookableResourceResponse();
            try
            {
                QueryExpression query = new QueryExpression("bookableresource");
                query.ColumnSet = new ColumnSet("msdyn_warehouse", "hil_defectivewarehouse");
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("userid", ConditionOperator.Equal, _userID.Id);
                EntityCollection _entitys = service.RetrieveMultiple(query);
                if (_entitys.Entities.Count == 1)
                {
                    bookableResource.BooableResource = _entitys[0].ToEntityReference();
                    bookableResource.DefectiveWarehouse = _entitys[0].GetAttributeValue<EntityReference>("hil_defectivewarehouse");
                    bookableResource.FreshWareHouse = _entitys[0].GetAttributeValue<EntityReference>("msdyn_warehouse");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error in Reteiving Bookable Resource. " + ex.Message);
            }
            return bookableResource;
        }
        public static Guid createPO(IOrganizationService service, PurchaseOrderModel purchaseOrderTypeModel)
        {
            Guid POID = Guid.Empty;
            EntityCollection JobPrdColl = null;
            try
            {
                QueryExpression query = new QueryExpression("msdyn_purchaseorder");
                query.ColumnSet = new ColumnSet(false);
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, purchaseOrderTypeModel.msdyn_workorder.Id);
                query.Criteria.AddCondition("hil_sapsalesorderno", ConditionOperator.Null);
                query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                EntityCollection _entitys = service.RetrieveMultiple(query);
                if (_entitys.Entities.Count != 0)
                {
                    POID = _entitys.Entities[0].Id;
                    JobPrdColl = RetrivePOProductToCreate(service, purchaseOrderTypeModel.msdyn_workorder, POID);
                }
                else
                {
                    JobPrdColl = RetrivePOProductToCreate(service, purchaseOrderTypeModel.msdyn_workorder, POID);
                    if (JobPrdColl.Entities.Count > 0)
                    {
                        Entity entity = new Entity("msdyn_purchaseorder");
                        entity["hil_potype"] = purchaseOrderTypeModel.hil_potype;
                        entity["msdyn_vendor"] = purchaseOrderTypeModel.msdyn_vendor;
                        entity["msdyn_receivetowarehouse"] = purchaseOrderTypeModel.msdyn_receivetowarehouse;
                        entity["msdyn_purchaseorderdate"] = purchaseOrderTypeModel.msdyn_purchaseorderdate;
                        entity["msdyn_requestedbyresource"] = purchaseOrderTypeModel.msdyn_requestedbyresource;
                        entity["ownerid"] = purchaseOrderTypeModel.ownerid;
                        entity["msdyn_orderedby"] = purchaseOrderTypeModel.msdyn_orderedby;
                        entity["msdyn_workorder"] = purchaseOrderTypeModel.msdyn_workorder;
                        entity["msdyn_systemstatus"] = new OptionSetValue(690970000);
                        entity["msdyn_name"] = "dd";
                        entity["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("68a6a9ca-6beb-e811-a96c-000d3af05828"));
                        POID = service.Create(entity);
                    }
                }
                if (POID != Guid.Empty)
                    CreatePOProducts(service, new EntityReference("msdyn_purchaseorder", POID), JobPrdColl, purchaseOrderTypeModel);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            return Guid.Empty;
        }
        public static void CreatePOProducts(IOrganizationService service, EntityReference PORef, EntityCollection JobPrdColl, PurchaseOrderModel purchaseOrderTypeModel)
        {
            foreach (Entity item in JobPrdColl.Entities)
            {
                Entity _POProduct = new Entity("msdyn_purchaseorderproduct");
                _POProduct["msdyn_product"] = item["hil_replacedpart"];
                _POProduct["msdyn_quantity"] = item["msdyn_quantity"];
                _POProduct["msdyn_unit"] = (EntityReference)item.GetAttributeValue<AliasedValue>("prod.defaultuomid").Value;
                _POProduct["msdyn_purchaseorder"] = PORef;// new EntityReference("msdyn_purchaseorder", POID);
                _POProduct["msdyn_associatetowarehouse"] = purchaseOrderTypeModel.msdyn_receivetowarehouse;
                _POProduct["msdyn_associatetoworkorder"] = purchaseOrderTypeModel.msdyn_workorder;
                _POProduct["hil_associatetoworkorderproduct"] = item.ToEntityReference();
                service.Create(_POProduct);
            }
        }
        public static EntityCollection RetrivePOProductToCreate(IOrganizationService service, EntityReference _jobID, Guid POID)
        {
            EntityCollection _entityCollection = new EntityCollection();
            try
            {
                {
                    string fetch = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                  <entity name=""msdyn_workorderproduct"">
                                    <attribute name=""createdon"" />
                                    <attribute name=""msdyn_product"" />
                                    <attribute name=""msdyn_linestatus"" />
                                    <attribute name=""hil_replacedpart"" />
                                    <attribute name=""msdyn_quantity"" />
                                    <attribute name=""msdyn_workorderproductid"" />
                                    <order attribute=""msdyn_product"" descending=""false"" />
                                    <filter type=""and"">
                                      <condition attribute=""msdyn_workorder"" operator=""eq"" value=""{_jobID.Id}"" />
                                      <condition attribute=""hil_replacedpart"" operator=""not-null"" />
                                      <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                                      <condition attribute=""hil_availabilitystatus"" operator=""ne"" value=""1"" />
                                    </filter>
                                    <link-entity name=""product"" from=""productid"" to=""hil_replacedpart"" visible=""false"" link-type=""outer"" alias=""prod"">
                                      <attribute name=""defaultuomid"" />
                                    </link-entity>
                                  </entity>
                                </fetch>";

                    EntityCollection JobPrdColl = service.RetrieveMultiple(new FetchExpression(fetch));
                    foreach (Entity item in JobPrdColl.Entities)
                    {
                        QueryExpression query = new QueryExpression("msdyn_purchaseorderproduct");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_associatetoworkorderproduct", ConditionOperator.Equal, item.Id);
                        query.Criteria.AddCondition("msdyn_associatetoworkorder", ConditionOperator.Equal, _jobID.Id);
                        query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        EntityCollection _entitysPrd = service.RetrieveMultiple(query);
                        if (_entitysPrd.Entities.Count == 0)
                        {
                            _entityCollection.Entities.Add(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            return _entityCollection;
        }
        public static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;
        }


        public static SendPaymentUrlResponse SendSMS(SendURLD365Request reqParm, IOrganizationService service)
        {
            SendPaymentUrlResponse sendPaymentUrlResponse = new SendPaymentUrlResponse();
            try
            {
                SendPaymentUrlRequest req = new SendPaymentUrlRequest();
                String comm = "create_invoice";
                req.PROJECT = "D365";
                req.command = comm.Trim();
                RemotePaymentLinkDetails remotePaymentLinkDetails = new RemotePaymentLinkDetails();
                if (service != null)
                {
                    if (reqParm.jobId == null)
                    {
                        sendPaymentUrlResponse.StatusCode = "Invalid Job GUID";
                    }
                    else
                    {
                        Entity job = service.Retrieve("msdyn_workorder", new Guid(reqParm.jobId), new ColumnSet(true));

                        QueryExpression Query = new QueryExpression("hil_paymentstatus");
                        Query.ColumnSet = new ColumnSet("hil_url");
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, job.GetAttributeValue<string>("msdyn_name"));
                        EntityCollection Found = service.RetrieveMultiple(Query);
                        if (Found.Entities.Count > 0)
                        {
                            sendPaymentUrlResponse.StatusCode = "SMS/Email is Allready send to Customer";
                            sendPaymentUrlResponse.URL = Found[0].GetAttributeValue<string>("hil_url");
                        }

                        //decimal abc = job.GetAttributeValue<Money>("hil_receiptamount").Value;

                        String state = job.Contains("hil_state") ? job.GetAttributeValue<EntityReference>("hil_state").Name.ToString() : string.Empty;
                        String zip = string.Empty;

                        string address = job.Contains("hil_fulladdress") ? job.GetAttributeValue<String>("hil_fulladdress").ToString() : string.Empty;
                        zip = job.Contains("hil_pincode") ? job.GetAttributeValue<EntityReference>("hil_pincode").Name.ToString() : string.Empty;

                        //if (address != string.Empty)
                        //{
                        //    zip = address.Substring(address.Length - 6);
                        //}
                        string city = string.Empty;
                        //if (job.Contains("hil_city"))
                        //{
                        //    if (state.ToLower() == "delhi")
                        //    {
                        //        city = "Delhi";
                        //    }
                        //    else
                        //    {
                        //        String cit = job.GetAttributeValue<EntityReference>("hil_city").Name.ToString();
                        //        city = cit.Substring(0, cit.Length - 3);
                        //    }
                        //}
                        remotePaymentLinkDetails.amount = reqParm.Amount;

                        Entity ent = service.Retrieve("hil_branch", job.GetAttributeValue<EntityReference>("hil_branch").Id, new ColumnSet("hil_mamorandumcode"));
                        string _txnId = string.Empty;
                        string _mamorandumCode = "";
                        if (ent.Attributes.Contains("hil_mamorandumcode"))
                        {
                            _mamorandumCode = ent.GetAttributeValue<string>("hil_mamorandumcode");
                        }

                        _txnId = "D365_" + job.GetAttributeValue<string>("msdyn_name");
                        remotePaymentLinkDetails.txnid = _txnId;
                        remotePaymentLinkDetails.firstname = job.Contains("hil_customerref") ? job.GetAttributeValue<EntityReference>("hil_customerref").Name.ToString() : string.Empty;
                        remotePaymentLinkDetails.email = job.Contains("hil_email") ? job.GetAttributeValue<String>("hil_email").ToString() : "abc@gmail.com";
                        remotePaymentLinkDetails.phone = reqParm.mobile;
                        //remotePaymentLinkDetails.city = city;//job.Contains("hil_city") ? ((state.ToLower() == "delhi") ? "Delhi" : job.GetAttributeValue<EntityReference>("hil_city").Name.ToString()) : string.Empty;
                        remotePaymentLinkDetails.address1 = address.Length > 99 ? address.Substring(0, 99) : address;
                        remotePaymentLinkDetails.state = state;
                        remotePaymentLinkDetails.country = "India";
                        remotePaymentLinkDetails.template_id = "1";
                        remotePaymentLinkDetails.productinfo = _mamorandumCode; //"B2C_PAYUBIZ_TEST_SMS";
                        remotePaymentLinkDetails.validation_period = "24";
                        remotePaymentLinkDetails.send_email_now = "1";
                        remotePaymentLinkDetails.send_sms = "1";
                        remotePaymentLinkDetails.time_unit = "H";
                        remotePaymentLinkDetails.zipcode = zip;
                        req.RemotePaymentLinkDetails = remotePaymentLinkDetails;
                        var client = new RestClient("https://middleware.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI");
                        client.Timeout = -1;
                        var request = new RestRequest(Method.POST);
                        string authInfo = "D365_HAVELLS" + ":" + "PRDD365@1234";
                        authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
                        request.AddHeader("Authorization", authInfo);
                        //request.AddHeader("Authorization", "Basic RDM2NV9IQVZFTExTOlFBRDM2NUAxMjM0");
                        request.AddHeader("Content-Type", "application/json");
                        request.AddHeader("Cookie", "saplb_*=(J2EE2717920)2717950; JSESSIONID=7fOj-tgnbYBRVBihJMBX9THzyTG3dgH-eCkA_SAPa_yX9TL_PrH5RR_PrxfO7kbO");
                        request.AddParameter("application/json", Newtonsoft.Json.JsonConvert.SerializeObject(req), ParameterType.RequestBody);
                        IRestResponse response = client.Execute(request);

                        var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SendPaymentUrlResponse>(response.Content);
                        if (obj.msg == null)
                        {
                            string url = obj.URL;
                            string[] invoicenumber = url.Split('=');
                            Entity statusPayment = new Entity("hil_paymentstatus");
                            statusPayment["hil_name"] = obj.Transaction_Id;
                            statusPayment["hil_url"] = obj.URL;
                            statusPayment["hil_statussendurl"] = obj.Status;
                            statusPayment["hil_email_id"] = obj.Email_Id;
                            statusPayment["hil_phone"] = obj.Phone;
                            statusPayment["hil_invoiceid"] = invoicenumber[1];
                            statusPayment["hil_job"] = new EntityReference("msdyn_workorder", new Guid(reqParm.jobId));
                            service.Create(statusPayment);

                            #region Updating Job Payment Link Sent field
                            Entity _updateJob = new Entity("msdyn_workorder", new Guid(reqParm.jobId));
                            _updateJob["hil_paymentlinksent"] = true;
                            service.Update(_updateJob);
                            #endregion
                            sendPaymentUrlResponse.StatusCode = obj.Status;
                            sendPaymentUrlResponse.URL = obj.URL;
                        }
                        else
                            sendPaymentUrlResponse.StatusCode = obj.msg;
                    }
                }
            }
            catch (Exception ex)
            {
                sendPaymentUrlResponse.StatusCode = "D365 Internal Error " + ex.Message;
            }
            return sendPaymentUrlResponse;
        }

    }
    public class TransactionDetail
    {

        public string mihpayid { get; set; }

        public string request_id { get; set; }

        public string bank_ref_num { get; set; }

        public string amt { get; set; }

        public string transaction_amount { get; set; }

        public string txnid { get; set; }

        public string additional_charges { get; set; }

        public string productinfo { get; set; }

        public string firstname { get; set; }

        public string bankcode { get; set; }

        public string udf1 { get; set; }

        public string udf3 { get; set; }

        public string udf4 { get; set; }

        public string udf5 { get; set; }

        public string field2 { get; set; }

        public string field9 { get; set; }

        public string error_code { get; set; }

        public string addedon { get; set; }

        public string payment_source { get; set; }

        public string card_type { get; set; }

        public string error_Message { get; set; }

        public string net_amount_debit { get; set; }

        public string disc { get; set; }

        public string mode { get; set; }

        public string PG_TYPE { get; set; }

        public string card_no { get; set; }

        public string udf2 { get; set; }

        public string status { get; set; }

        public string unmappedstatus { get; set; }

        public string Merchant_UTR { get; set; }

        public string Settled_At { get; set; }
    }
    public class SendPaymentUrlResponse
    {

        public string Email_Id { get; set; }

        public string Transaction_Id { get; set; }

        public string URL { get; set; }

        public string Status { get; set; }

        public string Phone { get; set; }

        public string StatusCode { get; set; }

        public string msg { get; set; }
    }
    public class SendURLD365Request
    {

        public string jobId { get; set; }

        public string mobile { get; set; }

        public string Amount { get; set; }
    }
    public class PaymentStatusD365Response
    {
        public string Status { get; set; }
    }
    public class PurchaseOrderType
    {
        public static readonly Guid Emergency = new Guid("784405c2-974c-ee11-be6f-6045bdac526a");
        public static readonly Guid ManualOrder = new Guid("3da37e41-7557-ee11-be6f-6045bdac5292");
        public static readonly Guid MinimumStockLevel = new Guid("9c0a9f6c-9157-ee11-be6e-6045bdaa91c3");
    }
    public class PurchaseOrderModel
    {
        public EntityReference hil_potype { get; set; }
        public EntityReference msdyn_vendor { get; set; }
        public EntityReference msdyn_receivetowarehouse { get; set; }
        public DateTime msdyn_purchaseorderdate { get; set; }
        public EntityReference msdyn_requestedbyresource { get; set; }
        public EntityReference msdyn_orderedby { get; set; }
        public EntityReference ownerid { get; set; }
        public EntityReference msdyn_workorder { get; set; }

    }
    public class SendPaymentUrlRequest
    {

        public string PROJECT { get; set; }

        public string command { get; set; }

        public RemotePaymentLinkDetails RemotePaymentLinkDetails { get; set; }
    }
    public class RemotePaymentLinkDetails
    {

        public string amount { get; set; }

        public string txnid { get; set; }

        public string productinfo { get; set; }

        public string firstname { get; set; }

        public string email { get; set; }

        public string phone { get; set; }

        public string address1 { get; set; }

        public string city { get; set; }

        public string state { get; set; }

        public string country { get; set; }

        public string zipcode { get; set; }

        public string template_id { get; set; }

        public string validation_period { get; set; }

        public string send_email_now { get; set; }

        public string send_sms { get; set; }

        public string time_unit { get; set; }
    }

    public class RetriveBookableResourceResponse
    {
        public EntityReference BooableResource { get; set; }
        public EntityReference FreshWareHouse { get; set; }
        public EntityReference DefectiveWarehouse { get; set; }

    }
}
