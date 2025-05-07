using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Net;
using System.Text;

namespace AE01.Inventory
{
    internal class TestingInventory
    {
        private IOrganizationService service;
        public TestingInventory(IOrganizationService _service)
        {
            this.service = _service;
        }
        internal async Task ToCall()
        {

            UpdatePurchaseOrder(); 

            PostPurhaseOrderReceipt postPurhaseOrderReceiptObj = new PostPurhaseOrderReceipt(service);
            postPurhaseOrderReceiptObj.PostOrderReceipt();

            InventoryRMALine inventoryRMALineObj = new InventoryRMALine(service);
            inventoryRMALineObj.UpdateLineNumber();

            PostCreateUpdateJob postCreateUpdateJobObj = new PostCreateUpdateJob(service);
            postCreateUpdateJobObj.ToCall();

            Entity purchaseOrder = service.Retrieve("hil_inventorypurchaseorder", new Guid("42bc63az28-c239-ef11-a316-7c1e5205da90"), new ColumnSet("hil_ordertype", "hil_postatus", "hil_productdivision", "hil_salesoffice"));
            //purchaseOrder["hil_ordertype"] = new OptionSetValue(2);
            //purchaseOrder["hil_postatus"] = new OptionSetValue(2); 
            //service.Update(purchaseOrder); 


            OnPreCreate(purchaseOrder);
        }

        private void OnPreCreate(Entity purchaseOrder)
        {

            OptionSetValue orderType = purchaseOrder.Contains("hil_ordertype") ? purchaseOrder.GetAttributeValue<OptionSetValue>("hil_ordertype") : new OptionSetValue(-1);
            OptionSetValue orderStatus = purchaseOrder.Contains("hil_postatus") ? purchaseOrder.GetAttributeValue<OptionSetValue>("hil_postatus") : new OptionSetValue(-1);
            Guid callingUserId = new Guid("79636995-24ab-ec11-9840-6045bdad4f3b");

            if (orderType.Value == 1) // Emergency
            {
                Entity updatePurchaseOrder = new Entity(purchaseOrder.LogicalName, purchaseOrder.Id);
                updatePurchaseOrder["hil_approver"] = new EntityReference("systemuser", callingUserId);
                updatePurchaseOrder["hil_approvedby"] = new EntityReference("systemuser", callingUserId);
                updatePurchaseOrder["hil_approvedon"] = DateTime.Now;
                updatePurchaseOrder["hil_postatus"] = new OptionSetValue(3); //Approved 
                service.Update(updatePurchaseOrder);
            }
            else if (orderStatus.Value == 2)// Submitted for Approval
            {
                EntityReference productDivision = purchaseOrder.Contains("hil_productdivision") ? purchaseOrder.GetAttributeValue<EntityReference>("hil_productdivision") : null;
                EntityReference salesOffice = purchaseOrder.Contains("hil_salesoffice") ? purchaseOrder.GetAttributeValue<EntityReference>("hil_salesoffice") : null;

                if (productDivision != null && salesOffice != null)
                {
                    EntityReference branchHead = GetApprover(productDivision, salesOffice);
                    if (branchHead != null)
                    {
                        Entity updatePurchaseOrder = new Entity(purchaseOrder.LogicalName, purchaseOrder.Id);
                        updatePurchaseOrder["hil_approver"] = branchHead;
                        updatePurchaseOrder["hil_approvedby"] = branchHead;
                        service.Update(updatePurchaseOrder);
                    }
                }
            }

        }

        private EntityReference GetApprover(EntityReference productDivision, EntityReference salesOffice)
        {
            QueryExpression expression = new QueryExpression("hil_sbubranchmapping");
            expression.ColumnSet = new ColumnSet("hil_branchheaduser");
            expression.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            expression.Criteria.AddCondition("hil_productdivision", ConditionOperator.Equal, productDivision.Id);
            expression.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, salesOffice.Id);

            EntityCollection approversColl = service.RetrieveMultiple(expression);
            if (approversColl.Entities.Count > 0)
            {
                return approversColl[0].GetAttributeValue<EntityReference>("hil_branchheaduser");
            }
            return null;

        }

        private void UpdatePurchaseOrder()
        {   //This function updates the channel partner on the purchase orders
            QueryExpression query = new QueryExpression("hil_inventorypurchaseorder");
            query.ColumnSet = new ColumnSet("hil_channelpartnercode", "hil_franchise", "hil_name");
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true; 

            EntityCollection purchaseOrders = service.RetrieveMultiple(query);

            EntityCollection finalColl = new EntityCollection(); 
            foreach(Entity i in purchaseOrders.Entities)
            {
                finalColl.Entities.Add(i);
            }
            do
            {
                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = purchaseOrders.PagingCookie;
                purchaseOrders = service.RetrieveMultiple(query);
                foreach (Entity i in purchaseOrders.Entities)
                {
                    finalColl.Entities.Add(i);

                }

            } while (purchaseOrders.MoreRecords);
            
            Console.WriteLine("Total purchase orders found are " + finalColl.Entities.Count); 

            for(int i = 0; i<finalColl.Entities.Count; i++)
            {
                Console.WriteLine("Processsing the order: " + (i +1));
                Console.WriteLine("Processing the order with number " + finalColl.Entities[i].GetAttributeValue<string>("hil_name"));

                EntityReference franchiseRef = finalColl.Entities[i].GetAttributeValue<EntityReference>("hil_franchise");

                if (franchiseRef != null)
                {
                    Entity franchise = service.Retrieve(franchiseRef.LogicalName, franchiseRef.Id, new ColumnSet("accountnumber"));
                    string accountNumber = franchise.GetAttributeValue<string>("accountnumber");
                    Entity updateChannelPartner = new Entity(finalColl.Entities[i].LogicalName, finalColl.Entities[i].Id);
                    updateChannelPartner["hil_channelpartnercode"] = accountNumber;
                    service.Update(updateChannelPartner);

                }
                
            }

        }
    }

    class SAPAPI
    {
        public class LTINVOICE
        {
            public string VBELN { get; set; }
            public string POSNR { get; set; }
            public decimal ECGST_PERC { get; set; }
            public string FKART { get; set; }
            public DateTime FKDAT { get; set; }
            public string AUBEL { get; set; }
            public string NETWR { get; set; }
            public string MWSBK { get; set; }
            public string TOT_DISC { get; set; }
            public string KUNRG { get; set; }
            public string WERKS { get; set; }
            public string BUTXT { get; set; }
            public string BEZEI { get; set; }
            public string MATNR { get; set; }
            public string FKIMG { get; set; }
            public string MEINS { get; set; }
            public string CHARG { get; set; }
            public string MRP { get; set; }
            public string DLP { get; set; }
            public string PUR_RATE { get; set; }
            public string LLAMOUNT { get; set; }
            public string LLDISCOUNT { get; set; }
            public string LLTAXAMT { get; set; }
            public string LLTOTAMT { get; set; }
            public string BT_MANUF_DATE { get; set; }
            public string BT_EXP_DATE { get; set; }
            public string ROUNDOFF { get; set; }
            public string FKSTO { get; set; }
            public string TAXABLE_AMT { get; set; }
            public string CGST_PERC { get; set; }
            public string SGST_PERC { get; set; }
            public decimal UTGST_PERC { get; set; }
            public string IGST_PERC { get; set; }
            public string CESS_PERC { get; set; }
            public string CGST_AMT { get; set; }
            public string SGST_AMT { get; set; }
            public decimal UTGST_AMT { get; set; }
            public string IGST_AMT { get; set; }
            public int STEUC { get; set; }
            public string REL_PARTY { get; set; }
            public string INV_TYPE { get; set; }
        }

        public class InvoiceResonse
        {
            public List<LTINVOICE> LT_INVOICE { get; set; }
        }

        class Request
        {
            public string ZREF_D365 { get; set; }
            public string KUNNR { get; set; }
            public string RETURN_TYPE { get; set; }
            public string REF_DATE { get; set; }
            public List<RmaLine> LT_TABLE;
        }

        public class Response
        {
            public string ZREF_D365 { get; set; }
            public string INSPNO { get; set; }
            public string ERROR { get; set; }
            public string STATUS { get; set; }
        }

        class RmaLine
        {
            public string MATNR { get; set; }
            public int RNQTY { get; set; }
        }

        private IOrganizationService service;

        public SAPAPI(IOrganizationService _service)
        {
            this.service = _service;
        }

        public class HelperClass
        {
            public static IntegrationConfiguration GetIntegrationConfiguration(IOrganizationService _service, string APIName)
            {
                try
                {
                    IntegrationConfiguration inconfig = new IntegrationConfiguration();
                    QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                    qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                    qsCType.NoLock = true;
                    qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                    qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, APIName);
                    Entity integrationConfiguration = _service.RetrieveMultiple(qsCType)[0];
                    inconfig.url = integrationConfiguration.GetAttributeValue<string>("hil_url");
                    inconfig.userName = integrationConfiguration.GetAttributeValue<string>("hil_username");
                    inconfig.password = integrationConfiguration.GetAttributeValue<string>("hil_password");
                    return inconfig;
                }
                catch (Exception ex)
                {
                    throw new Exception("Error : " + ex.Message);
                }
            }
            public static string CallAPI(IntegrationConfiguration integrationConfiguration, string Json, String method)
            {
                WebRequest request = WebRequest.Create(integrationConfiguration.url);
                request.Headers[HttpRequestHeader.Authorization] = "Basic " + Convert.ToBase64String(Encoding.ASCII.GetBytes(integrationConfiguration.userName + ":" + integrationConfiguration.password));
                request.Method = method; //"POST";
                if (!string.IsNullOrEmpty(Json))
                {
                    byte[] byteArray = Encoding.UTF8.GetBytes(Json);
                    request.ContentType = "application/x-www-form-urlencoded";
                    request.ContentLength = byteArray.Length;
                    Stream dataStream = request.GetRequestStream();
                    dataStream.Write(byteArray, 0, byteArray.Length);
                    dataStream.Close();
                }
                WebResponse response = request.GetResponse();
                Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                Stream dataStream1 = response.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream1);
                return reader.ReadToEnd();
            }
        }

        internal async Task PostRMA()
        {
            QueryExpression queryExpression = new QueryExpression("hil_inventoryrma");
            queryExpression.ColumnSet = new ColumnSet("hil_name", "hil_franchise", "hil_returntype", "hil_responsefromsap", "hil_syncedwithsap", "hil_inspectionnumber", "hil_rmastatus");
            queryExpression.Criteria.AddCondition("hil_rmastatus", ConditionOperator.Equal, 4); // Posted
            queryExpression.Criteria.AddCondition("hil_syncedwithsap", ConditionOperator.Equal, false);
            EntityCollection rmaCollection = service.RetrieveMultiple(queryExpression);

            foreach (Entity inventoryRMA in rmaCollection.Entities)
            {


                Request request = new Request();
                request.REF_DATE = DateTime.Now.Date.ToString("yyyy-MM-dd").Replace("-", "");
                request.ZREF_D365 = inventoryRMA.Contains("hil_name") ? inventoryRMA.GetAttributeValue<string>("hil_name") : string.Empty;


                EntityReference franchiseRef = inventoryRMA.GetAttributeValue<EntityReference>("hil_franchise");
                Entity franchise = service.Retrieve("account", franchiseRef.Id, new ColumnSet("accountnumber"));
                request.KUNNR = "F" + franchise.GetAttributeValue<string>("accountnumber");
                Console.WriteLine("Franchise Code: " + request.KUNNR);

                EntityReference rmaTypeRef = inventoryRMA.GetAttributeValue<EntityReference>("hil_returntype");
                Entity rmaType = service.Retrieve(rmaTypeRef.LogicalName, rmaTypeRef.Id, new ColumnSet("hil_sapcode"));
                request.RETURN_TYPE = rmaType.GetAttributeValue<string>("hil_sapcode");

                QueryExpression query = new QueryExpression("hil_inventoryrmaline");
                query.ColumnSet = new ColumnSet("hil_product", "hil_quantity");
                query.Criteria.AddCondition("hil_rma", ConditionOperator.Equal, inventoryRMA.Id);
                EntityCollection linesCollection = service.RetrieveMultiple(query);
                List<RmaLine> tempList = new List<RmaLine>();

                for (int i = 0; i < linesCollection.Entities.Count; i++)
                {
                    RmaLine tempLine = new RmaLine();
                    EntityReference productRef = linesCollection.Entities[i].GetAttributeValue<EntityReference>("hil_product");
                    Entity product = service.Retrieve(productRef.LogicalName, productRef.Id, new ColumnSet("productnumber"));
                    tempLine.MATNR = product.Contains("productnumber") ? product.GetAttributeValue<string>("productnumber") : string.Empty;
                    tempLine.RNQTY = linesCollection.Entities[i].Contains("hil_quantity") ? linesCollection.Entities[i].GetAttributeValue<int>("hil_quantity") : 0;
                    tempList.Add(tempLine);
                }
                request.LT_TABLE = tempList;

                var client = new HttpClient();
                var apiRequest = new HttpRequestMessage(HttpMethod.Post, "https://p90ci.havells.com:50001/RESTAdapter/d365/franchisee_return/insp_creation");
                apiRequest.Headers.Add("Authorization", "Basic RDM2NV9IQVZFTExTOlBSREQzNjVAMTIzNA==");
                HttpContent content = new StringContent(JsonConvert.SerializeObject(request));
                apiRequest.Content = content;
                var response = await client.SendAsync(apiRequest);
                response.EnsureSuccessStatusCode();
                Response res = JsonConvert.DeserializeObject<Response>(await response.Content.ReadAsStringAsync());

                if (res != null)
                {
                    inventoryRMA["hil_responsefromsap"] = res.ERROR;
                    if (res.STATUS == "S")
                    {
                        inventoryRMA["hil_syncedwithsap"] = true;
                        inventoryRMA["hil_responsefromsap"] = "Success";
                    }
                    if (!string.IsNullOrEmpty(res.INSPNO)) inventoryRMA["hil_inspectionnumber"] = res.INSPNO;
                    service.Update(inventoryRMA);
                }
            }

        }

        public class PO_SAP_Request
        {
            public string IM_PROJECT { get; set; }
            public List<IT_DATA> LT_LINE_ITEM { get; set; }
            public Parent IM_HEADER { get; set; }
        }

        public class IT_DATA
        {
            public string MATNR { get; set; }
            public int DZMENG { get; set; }
            public string SPART { get; set; }
        }

        public class IMPROJECT
        {
            public string IM_PROJECT { get; set; }
        }

        public class Parent
        {
            public string AUART { get; set; }
            public string VTWEG { get; set; }
            public string SPART { get; set; }
            public string BSTKD { get; set; }
            public string BSTDK { get; set; }
            public string KUNNR { get; set; }
            public string ABRVW { get; set; }
            public string VKORG { get; set; }
        }

        public class IntegrationLog
        {
            public Guid hil_integrationlogid { get; set; }
            public string hil_name { get; set; }
            public int hil_totalrecords { get; set; }
            public int hil_recordsaffected { get; set; }
            public int hil_totalerrorrecords { get; set; }
            public bool hil_states { get; set; }
            public string hil_syncdate { get; set; }
        }

        public class Response_SAP
        {
            public string EX_SALESDOC_NO { get; set; }
            public string RETURN { get; set; }
            public List<ET_Details> ET_SO_DETAILS;
        }

        public class ET_Details
        {
            public string VBELN { get; set; }
            public string POSNR { get; set; }
            public string MATNR { get; set; }
            public string MAKTX { get; set; }
            public string ERDAT { get; set; }
            public string KWMENG { get; set; }
            public string VRKME { get; set; }
            public string NETWR { get; set; }
            public string WAERK { get; set; }
            public string NETPR { get; set; }
            public string MWSBP { get; set; }
            public string H_LFSTK { get; set; }
            public string H_STATUS { get; set; }
            public string L_LFSTA { get; set; }
            public string L_STATUS { get; set; }
            public string KUNNR { get; set; }
            public string H_DEL_DATE { get; set; }
            public string L_DEL_DATE { get; set; }
            public string BSTNK { get; set; }
            public string IHREZ { get; set; }
            public string H_NETWR { get; set; }
            public string H_WAERK { get; set; }
            public string ST_KUNNR { get; set; }
            public string NAME1 { get; set; }
            public string NAME2 { get; set; }
            public string LAND1 { get; set; }
            public string ORT01 { get; set; }
            public string PSTLZ { get; set; }
            public string REGIO { get; set; }
            public string STRAS { get; set; }
            public string TELF1 { get; set; }
            public string TELFX { get; set; }
            public string REGSMTP_ADDRIO { get; set; }
        }


        internal void SyncPoToSAP()
        {
            string Date = string.Empty;
            int i = 0;
            string sUserName = "D365_HAVELLS";
            string sPassword = "PRDD365@1234";

            QueryExpression Query = new QueryExpression("hil_inventorypurchaseorder");
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_postatus", ConditionOperator.Equal, 3)); //Approved 
            // Check for the active purchase orders.
            Query.AddOrder("modifiedon", OrderType.Descending);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                //HeaderId = Found.Entities[0].Id;
                PO_SAP_Request SAPrequest = new PO_SAP_Request();
                IntegrationLog il = new IntegrationLog();

                //Guid integrationHeaderId = IntegrationLogUtils.CreateLogHeader(service, il);

                foreach (Entity purchaseOrderP in Found.Entities)
                {
                    try
                    {

                        //->// Remove the line 
                        Entity purchaseOrder = service.Retrieve("hil_inventorypurchaseorder", new Guid("4684833d-9254-ef11-bfe2-6045bdc665f4"), new ColumnSet(true));
                        // if(true)
                        i = Found.Entities.Count;
                        if (purchaseOrder.Contains("createdon"))
                        {
                            DateTime iDate = purchaseOrder.GetAttributeValue<DateTime>("createdon");
                            Date = iDate.AddMinutes(330).ToString("yyyy-MM-dd");
                        }

                        EntityReference productRef = purchaseOrder.GetAttributeValue<EntityReference>("hil_productdivision");
                        Entity product = service.Retrieve(productRef.LogicalName, productRef.Id, new ColumnSet("productnumber", "name"));

                        EntityReference franchiseRef = purchaseOrder.GetAttributeValue<EntityReference>("hil_franchise");
                        Entity franchise = service.Retrieve("account", franchiseRef.Id, new ColumnSet("accountnumber"));

                        EntityReference divRef = purchaseOrder.GetAttributeValue<EntityReference>("hil_productdivision");
                        Entity div = service.Retrieve(divRef.LogicalName, divRef.Id, new ColumnSet("hil_sapcode"));
                        string SAPCode = div.GetAttributeValue<string>("hil_sapcode");


                        Parent iParent = new Parent();
                        iParent.SPART = SAPCode;//product.Contains("productnumber") ? product.GetAttributeValue<string>("productnumber") : string.Empty;
                        iParent.BSTKD = purchaseOrder.GetAttributeValue<string>("hil_name");


                        iParent.BSTDK = Date;
                        iParent.KUNNR = "F" + franchise.GetAttributeValue<string>("accountnumber");
                        iParent.VKORG = "HIL";
                        iParent.ABRVW = "EM";
                        iParent.AUART = "ZRS4";
                        iParent.VTWEG = "56";


                        IMPROJECT project = new IMPROJECT();
                        project.IM_PROJECT = "D365";

                        QueryExpression Query1 = new QueryExpression("hil_inventorypurchaseorderline");
                        Query1.ColumnSet = new ColumnSet(true);
                        Query1.Criteria = new FilterExpression(LogicalOperator.And);
                        Query1.Criteria.AddCondition(new ConditionExpression("hil_ponumber", ConditionOperator.Equal, purchaseOrder.Id));
                        Query1.Criteria.AddCondition(new ConditionExpression("hil_orderquantity", ConditionOperator.NotNull));
                        EntityCollection Found1 = service.RetrieveMultiple(Query1);
                        if (Found1.Entities.Count > 0)
                        {
                            List<IT_DATA> childs = new List<IT_DATA>();
                            foreach (Entity orderLine in Found1.Entities)
                            {
                                IT_DATA iChild = new IT_DATA();
                                iChild.MATNR = orderLine.GetAttributeValue<EntityReference>("hil_partcode").Name;
                                iChild.SPART = SAPCode;
                                iChild.DZMENG = orderLine.GetAttributeValue<int>("hil_orderquantity");
                                childs.Add(iChild);
                            }
                            SAPrequest.IM_PROJECT = "D365";
                            SAPrequest.IM_HEADER = iParent;
                            SAPrequest.LT_LINE_ITEM = childs;
                            var Json = JsonConvert.SerializeObject(SAPrequest);
                            //Production URL
                            WebRequest request = WebRequest.Create("https://p90ci.havells.com:50001/RESTAdapter/Common/CreateSalesOrder");
                            //WebRequest request = WebRequest.Create("https://middleware.havells.com:50001/RESTAdapter/Common/CreateSalesOrder");
                            //UAT URL
                            //WebRequest request = WebRequest.Create("https://middlewareqa.havells.com:50001/RESTAdapter/Common/CreateSalesOrder");
                            string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
                            request.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                            // Set the Method property of the request to POST.  
                            request.Method = "POST";
                            byte[] byteArray = Encoding.UTF8.GetBytes(Json);
                            // Set the ContentType property of the WebRequest.  
                            request.ContentType = "application/x-www-form-urlencoded";
                            // Set the ContentLength property of the WebRequest.  
                            request.ContentLength = byteArray.Length;
                            // Get the request stream.  
                            Stream dataStream = request.GetRequestStream();
                            // Write the data to the request stream.  
                            dataStream.Write(byteArray, 0, byteArray.Length);
                            // Close the Stream object.  
                            dataStream.Close();
                            // Get the response.  
                            WebResponse response = request.GetResponse();
                            // Display the status.  
                            Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                            // Get the stream containing content returned by the server.  
                            dataStream = response.GetResponseStream();
                            // Open the stream using a StreamReader for easy access.  
                            StreamReader reader = new StreamReader(dataStream);
                            // Read the content.  
                            string responseFromServer = reader.ReadToEnd();
                            // responseFromServer = responseFromServer.Replace("ZBAPI_CREATE_SALES_ORDER.Response", "ZBAPI_CREATE_SALES_ORDER"); // removed for new response
                            Response_SAP resp = JsonConvert.DeserializeObject<Response_SAP>(responseFromServer);
                            Console.WriteLine(resp.EX_SALESDOC_NO + " - " + resp.RETURN + " - " + purchaseOrder.GetAttributeValue<string>("hil_name"));
                            if (resp.EX_SALESDOC_NO != "")
                            {

                                purchaseOrder["hil_syncedwithsap"] = true;
                                purchaseOrder["hil_responsefromsap"] = resp.RETURN;
                                purchaseOrder["hil_salesordernumber"] = resp.EX_SALESDOC_NO.ToString();
                                service.Update(purchaseOrder);
                            }
                            else
                            {
                                purchaseOrder["hil_syncedwithsap"] = true;
                                purchaseOrder["hil_responsefromsap"] = resp.RETURN;
                                purchaseOrder["hil_salesordernumber"] = resp.EX_SALESDOC_NO.ToString();

                                service.Update(purchaseOrder);


                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
        }

        private Guid GetInventoryPurchaseRecept(string sapInvoiceNumber, EntityReference franchise, EntityReference purchaseOrder, EntityReference freshWarehouse, Guid ownerId)
        {
            QueryExpression query = new QueryExpression("hil_inventorypurchaseorderreceipt");
            query.Criteria.AddCondition("hil_invoicenumber", ConditionOperator.Equal, sapInvoiceNumber);
            EntityCollection receiptCollection = service.RetrieveMultiple(query);
            if (receiptCollection.Entities.Count > 0)
            {
                return receiptCollection.Entities[0].Id;
            }

            Entity newPurchaseReceipt = new Entity("hil_inventorypurchaseorderreceipt");
            newPurchaseReceipt["hil_franchise"] = franchise;
            newPurchaseReceipt["hil_invoicenumber"] = sapInvoiceNumber;
            newPurchaseReceipt["hil_ordernumber"] = purchaseOrder;
            newPurchaseReceipt["hil_warehouse"] = freshWarehouse;
            newPurchaseReceipt["hil_receiptstatus"] = new OptionSetValue(1);
            newPurchaseReceipt["ownerid"] = new EntityReference("systemuser", ownerId);
            return service.Create(newPurchaseReceipt);

        }

        private Guid GetPurchaseOrderReceiptLine(Guid purchaseOrderReceipt, EntityReference product, EntityReference purchaseOrder, EntityReference freshWarehouse, Guid ownerId)
        {
            QueryExpression query = new QueryExpression("hil_inventorypurchaseorderreceiptline");
            query.Criteria.AddCondition("hil_receiptnumber", ConditionOperator.Equal, purchaseOrderReceipt);
            query.Criteria.AddCondition("hil_partcode", ConditionOperator.Equal, product.Id);
            EntityCollection linesCollection = service.RetrieveMultiple(query);
            if (linesCollection.Entities.Count > 0)
            {
                return linesCollection.Entities[0].Id;
            }

            Entity purchaseOrderLine = GetOrderLine(product, purchaseOrder);

            Entity newPurchaseLine = new Entity("hil_inventorypurchaseorderreceiptline");
            newPurchaseLine["hil_receiptnumber"] = new EntityReference("hil_inventorypurchaseorderreceipt", purchaseOrderReceipt);
            newPurchaseLine["hil_partcode"] = product;
            newPurchaseLine["hil_ordernumber"] = purchaseOrder;
            newPurchaseLine["hil_warehouse"] = freshWarehouse;
            newPurchaseLine["hil_purchaseorderline"] = purchaseOrderLine.ToEntityReference();
            newPurchaseLine["ownerid"] = new EntityReference("systemuser", ownerId);
            return service.Create(newPurchaseLine);

        }

        private Entity GetOrderLine(EntityReference product, EntityReference purhaseOrder)
        {
            QueryExpression query = new QueryExpression("hil_inventorypurchaseorderline");
            query.Criteria.AddCondition("hil_partcode", ConditionOperator.Equal, product.Id);
            query.Criteria.AddCondition("hil_ponumber", ConditionOperator.Equal, purhaseOrder.Id);

            EntityCollection orderLineColl = service.RetrieveMultiple(query);
            return orderLineColl.Entities[0];
        }

        internal void GetSapInvoicetoSync()
        {
            IntegrationConfiguration integrationConfiguration = HelperClass.GetIntegrationConfiguration(service, "CRM_SAPtoCRM_Invoice");
            DateTime FromDate = DateTime.Now.AddDays(-1);  //GetLastRunDate(service);
            String fromDate = FromDate.Year.ToString() + FromDate.Month.ToString().PadLeft(2, '0') + FromDate.Day.ToString().PadLeft(2, '0');

            DateTime ToDate = DateTime.Now;
            String toDate = ToDate.Year.ToString() + ToDate.Month.ToString().PadLeft(2, '0') + ToDate.Day.ToString().PadLeft(2, '0');

            //integrationConfiguration.url = "https://middlewareqa.havells.com:50001/RESTAdapter/dynamics/invoice?IM_FLAG=R&IM_PROJECT=SER";
            integrationConfiguration.url = integrationConfiguration.url.Replace("middleware", "p90ci");
            integrationConfiguration.url = integrationConfiguration.url + "&IM_FROM_DT=" + fromDate + "&IM_TO_DT=" + toDate;
            //integrationConfiguration.url = integrationConfiguration.url + "&IM_FROM_DT=" + "20240110" + "&IM_TO_DT=" + "20240130";
            //integrationConfiguration.url = integrationConfiguration.url.Replace("middleware", "p90ci");
            //integrationConfiguration.url = integrationConfiguration.url.Replace("middleware", "p90ci");

            //integrationConfiguration.password = "QAD365@1234";

            InvoiceResonse resp = JsonConvert.DeserializeObject<InvoiceResonse>(HelperClass.CallAPI(integrationConfiguration, null, "GET"));

            if (resp.LT_INVOICE.Count > 0)
            {

                QueryExpression query = new QueryExpression("hil_inventorypurchaseorder");
                query.ColumnSet = new ColumnSet("hil_salesordernumber");
                query.Criteria.AddCondition("createdon", ConditionOperator.Last7Days);
                EntityCollection poCollection = service.RetrieveMultiple(query);

                HashSet<string> salesOrderNumbers = new HashSet<string>();

                for (int i = 0; i < poCollection.Entities.Count; i++)
                {
                    if (poCollection.Entities[i].Contains("hil_salesordernumber"))
                    {
                        salesOrderNumbers.Add(poCollection.Entities[i].GetAttributeValue<string>("hil_salesordernumber"));
                    }
                }


                foreach (LTINVOICE lTINVOICE in resp.LT_INVOICE)
                {
                    try
                    {
                        //->// remove the order number from the below code 
                        if (lTINVOICE != null && salesOrderNumbers.Contains(lTINVOICE.AUBEL))
                        {
                            string franchiseCode = lTINVOICE.KUNRG.Substring(1);

                            EntityReference product = GetProduct(lTINVOICE.MATNR); // lTINVOICE.MATNR
                            EntityReference franchise = GetFranchise(franchiseCode); // lTINVOICE.KUNRG 

                            if (product != null && franchise != null)
                            {
                                query = new QueryExpression("hil_inventorypurchaseorder");
                                query.ColumnSet = new ColumnSet("ownerid");
                                query.Criteria.AddCondition("hil_salesordernumber", ConditionOperator.Equal, lTINVOICE.AUBEL);
                                EntityCollection purchaseOrderCollection = service.RetrieveMultiple(query);

                                query = new QueryExpression("hil_inventorywarehouse");
                                query.Criteria.AddCondition("hil_franchise", ConditionOperator.Equal, franchise.Id);
                                query.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 1); //fresh 
                                EntityCollection wareHouseCollection = service.RetrieveMultiple(query);

                                if (purchaseOrderCollection.Entities.Count > 0 && wareHouseCollection.Entities.Count > 0)
                                {
                                    Entity newInventorySpareBills = new Entity("hil_inventorysparebills");
                                    newInventorySpareBills["hil_salesordernumber"] = lTINVOICE.AUBEL;
                                    newInventorySpareBills["hil_billnumber"] = lTINVOICE.VBELN;
                                    newInventorySpareBills["hil_billdate"] = lTINVOICE.FKDAT;
                                    if (!string.IsNullOrEmpty(lTINVOICE.FKIMG))
                                    {
                                        newInventorySpareBills["hil_quantity"] = Convert.ToInt32(Convert.ToDecimal(lTINVOICE.FKIMG));//Quantity
                                    }
                                    newInventorySpareBills["hil_franchise"] = franchise;
                                    newInventorySpareBills["hil_partcode"] = product;

                                    if (purchaseOrderCollection.Entities.Count > 0)
                                    {
                                        newInventorySpareBills["hil_ordernumber"] = purchaseOrderCollection.Entities[0].ToEntityReference();
                                    }

                                    service.Create(newInventorySpareBills);

                                    EntityReference poOwner = purchaseOrderCollection.Entities[0].GetAttributeValue<EntityReference>("ownerid");

                                    Guid purchaseOrderReceipt = GetInventoryPurchaseRecept(lTINVOICE.VBELN, franchise, purchaseOrderCollection.Entities[0].ToEntityReference(), wareHouseCollection.Entities[0].ToEntityReference(), poOwner.Id);

                                    Guid purchaseOrderReceiptLine = GetPurchaseOrderReceiptLine(purchaseOrderReceipt, product, purchaseOrderCollection.Entities[0].ToEntityReference(), wareHouseCollection.Entities[0].ToEntityReference(), poOwner.Id);

                                    Entity updateLine = new Entity("hil_inventorypurchaseorderreceiptline", purchaseOrderReceiptLine);

                                    if (!string.IsNullOrEmpty(lTINVOICE.FKIMG))
                                    {
                                        updateLine["hil_billedquantity"] = Convert.ToInt32(Convert.ToDecimal(lTINVOICE.FKIMG));//Quantity
                                    }
                                    service.Update(updateLine);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("error " + ex.Message);
                    }
                }
            }


        }

        internal void CreateInventoryBill()
        {

            EntityReference product = GetProduct("GSCPGMHAX294"); // lTINVOICE.MATNR
            EntityReference franchise = GetFranchise("9425390647"); // lTINVOICE.KUNRG

            if (product != null && franchise != null)
            {
                Entity newInventorySpareBills = new Entity("hil_inventorysparebills");
                newInventorySpareBills["hil_salesordernumber"] = "HG987";//lTINVOICE.AUBEL;
                newInventorySpareBills["hil_billnumber"] = "9876";//lTINVOICE.VBELNl;
                newInventorySpareBills["hil_billdate"] = DateTime.Now; //lTINVOICE.FKDAT;
                //if (!string.IsNullOrEmpty(lTINVOICE.FKIMG))
                //{
                //    invoiceDeatils["hil_quantity"] = Convert.ToInt32(Convert.ToDecimal(lTINVOICE.FKIMG));//Quantity
                //}
                newInventorySpareBills["hil_quantity"] = 7;//
                newInventorySpareBills["hil_franchise"] = franchise;
                newInventorySpareBills["hil_partcode"] = product;
                service.Create(newInventorySpareBills);
            }
        }

        private EntityReference? GetFranchise(string franchiseCode)
        {
            QueryExpression query = new QueryExpression("account");
            query.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, franchiseCode);
            EntityCollection tempCollection = service.RetrieveMultiple(query);
            if (tempCollection.Entities.Count > 0)
            {
                return tempCollection.Entities[0].ToEntityReference();
            }
            return null;
        }

        private EntityReference? GetProduct(string productCode)
        {
            QueryExpression query = new QueryExpression("product");
            query.Criteria.AddCondition("hil_productcode", ConditionOperator.Equal, productCode);
            EntityCollection tempCollection = service.RetrieveMultiple(query);
            if (tempCollection.Entities.Count > 0)
            {
                return tempCollection.Entities[0].ToEntityReference();
            }
            return null;
        }

    }
    
    internal class PostCreateUpdateJob
    {
        private IOrganizationService service;

        private readonly EntityReference workInitiated = new EntityReference("msdyn_workordersubstatus", new Guid("2b27fa6c-fa0f-e911-a94e-000d3af060a1"));
    
        private readonly Guid workDone = new Guid("2927fa6c-fa0f-e911-a94e-000d3af060a1"); //workDoneSubStauts

        public PostCreateUpdateJob(IOrganizationService _service)
        {
            this.service = _service;
        }

        internal void ToCall()
        {  
            JobData data = new JobData();
            data.WorkOrder = service.Retrieve("msdyn_workorder", new Guid("98ffcaee-b675-ef11-ac21-6045bdad7f34"), new ColumnSet("hil_flagpo", "msdyn_substatus", "hil_owneraccount", "hil_productcategory", "hil_salesoffice", "hil_brand"));
            data.Franchise = data.WorkOrder.GetAttributeValue<EntityReference>("hil_owneraccount");
            Entity franchise = service.Retrieve(data.Franchise.LogicalName, data.Franchise.Id, new ColumnSet("ownerid", "hil_spareinventoryenabled"));
            bool isInventoryEnabled = franchise.GetAttributeValue<Boolean>("hil_spareinventoryenabled");

            bool flagPO = data.WorkOrder.GetAttributeValue<Boolean>("hil_flagpo");

            if (flagPO && isInventoryEnabled)
            {
                data.SubStatus = data.WorkOrder.GetAttributeValue<EntityReference>("msdyn_substatus");
                data.ProductCategory = data.WorkOrder.GetAttributeValue<EntityReference>("hil_productcategory");
                data.SalesOffice = data.WorkOrder.GetAttributeValue<EntityReference>("hil_salesoffice");
                data.Brand = data.WorkOrder.GetAttributeValue<OptionSetValue>("hil_brand");
                if (data.SubStatus.Id != workDone)
                {
                    string fetchProducts = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='msdyn_workorderproduct'>
                                            <attribute name='msdyn_product' />
                                            <attribute name='msdyn_workorderproductid' />
                                            <attribute name='hil_quantity' />
                                            <attribute name='msdyn_quantity' />
                                            <attribute name='msdyn_estimatequantity' />
                                            <order attribute='msdyn_product' descending='false' />
                                            <filter type='and'>
                                                <condition attribute='hil_replacedpart' operator='not-null' />
                                                <condition attribute='hil_availabilitystatus' operator='eq' value='2' />
                                                <condition attribute='msdyn_workorder' operator='eq' value='{data.WorkOrder.Id}' />
                                            </filter>
                                            <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' link-type='inner' alias='ad'>
                                                <attribute name='hil_owneraccount' />
                                                <attribute name='hil_productcategory' />
                                                <filter type='and'>
                                                <condition attribute='hil_flagpo' operator='eq' value='1' />
                                                </filter>
                                            </link-entity>
                                            </entity>
                                        </fetch>";

                    EntityCollection tempCollection = service.RetrieveMultiple(new FetchExpression(fetchProducts));
                    
                    if (tempCollection.Entities.Count > 0)
                    {
                        EntityReference franchiseOwner = franchise.Contains("ownerid") ? franchise.GetAttributeValue<EntityReference>("ownerid") : null;

                        Entity warehouse = GetFreshWarehouse(data.Franchise.Id);
                        if (warehouse != null)
                        {
                            // Create purchase order 
                            Guid purchaseOrder = DoesOrderExists(data, warehouse);
                            if (purchaseOrder == Guid.Empty)
                            {
                                purchaseOrder = CreatePurchaseOrder(data, warehouse, franchiseOwner);

                            }

                            // Create purchase order lines. 
                            foreach (var temp in tempCollection.Entities)
                            {
                                EntityReference product = temp.Contains("msdyn_product") ? temp.GetAttributeValue<EntityReference>("msdyn_product") : null;
                                double orderQuantity = temp.Contains("msdyn_quantity") ? temp.GetAttributeValue<double>("msdyn_quantity") : 0;
                                int warrantyStatus = temp.Contains("hil_warrantystatus") ? temp.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value : 1;
                                
                                if (!DoesOrderLineExists(temp, purchaseOrder, product, orderQuantity, warrantyStatus))
                                {
                                    CreateOrderLine(temp, purchaseOrder, product, orderQuantity, warrantyStatus, franchiseOwner);
                                }

                            }                            
                            //Update CRMAdmin as approver and approve the Purchase Order.
                            UpdateApproverOnPurchaseOrder(purchaseOrder);
                        }
                    }

                }
            }

        }
       
        private bool DoesOrderLineExists(Entity jobProduct, Guid purchaseOrder, EntityReference product, double orderQuantity, int warrantyStatus)
        {
            QueryExpression query = new QueryExpression("hil_inventorypurchaseorderline");
            query.Criteria.AddCondition("hil_jobproduct", ConditionOperator.Equal, jobProduct.Id);
            query.Criteria.AddCondition("hil_ponumber", ConditionOperator.Equal, purchaseOrder);
            query.Criteria.AddCondition("hil_partcode", ConditionOperator.Equal, product.Id);
            if (orderQuantity != -1) query.Criteria.AddCondition("hil_orderquantity", ConditionOperator.Equal, (int)orderQuantity);
            if (warrantyStatus != -1) query.Criteria.AddCondition("hil_warrantystatus", ConditionOperator.Equal, warrantyStatus);
            EntityCollection tempColl = service.RetrieveMultiple(query);
            if (tempColl.Entities.Count > 0)
            {
                return true;
            }
            return false;
        }

        private Guid CreateOrderLine(Entity jobProduct, Guid purchaseOrder, EntityReference product, double orderQuantity, int warrantyStatus, EntityReference owner)
        {
            Entity newPurchaseOrderLine = new Entity("hil_inventorypurchaseorderline");
            newPurchaseOrderLine["hil_jobproduct"] = jobProduct.ToEntityReference();
            newPurchaseOrderLine["hil_ponumber"] = new EntityReference("hil_inventorypurchaseorder", purchaseOrder);
            newPurchaseOrderLine["hil_partcode"] = product;
            newPurchaseOrderLine["ownerid"] = owner;
            if (orderQuantity != -1) newPurchaseOrderLine["hil_orderquantity"] = (int)orderQuantity;
            if (warrantyStatus != -1) newPurchaseOrderLine["hil_warrantystatus"] = new OptionSetValue(warrantyStatus);


            return service.Create(newPurchaseOrderLine);
        }

        private Guid DoesOrderExists(JobData data, Entity warehouse)
        {
            QueryExpression query = new QueryExpression("hil_inventorypurchaseorder");
            query.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, data.WorkOrder.Id);
            query.Criteria.AddCondition("hil_productdivision", ConditionOperator.Equal, data.ProductCategory.Id);
            query.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, data.SalesOffice.Id);
            query.Criteria.AddCondition("hil_franchise", ConditionOperator.Equal, data.Franchise.Id);
            query.Criteria.AddCondition("hil_warehouse", ConditionOperator.Equal, warehouse.Id);
            query.Criteria.AddCondition("hil_syncedwithsap", ConditionOperator.Equal, false);
            EntityCollection tempColl = service.RetrieveMultiple(query);
            if (tempColl.Entities.Count > 0)
            {
                return tempColl.Entities[0].Id;
            }
            return Guid.Empty;

        }

        private void UpdateApproverOnPurchaseOrder(Guid purchaseOrder)
        {
            Guid CrmAdmin = new Guid("5190416c-0782-e911-a959-000d3af06a98");
            Entity updatePurchaseOrder = new Entity("hil_inventorypurchaseorder", purchaseOrder);
            updatePurchaseOrder["hil_approver"] = new EntityReference("systemuser", CrmAdmin);
            updatePurchaseOrder["hil_approvedby"] = new EntityReference("systemuser", CrmAdmin);
            updatePurchaseOrder["hil_approvedon"] = DateTime.Now;
            updatePurchaseOrder["hil_postatus"] = new OptionSetValue(3); //Approved 
            service.Update(updatePurchaseOrder);
        }
       
        private Guid CreatePurchaseOrder(JobData data, Entity warehouse, EntityReference owner)
        {

           
            Entity newPurchaseOrder = new Entity("hil_inventorypurchaseorder");
            newPurchaseOrder["hil_jobid"] = new EntityReference(data.WorkOrder.LogicalName, data.WorkOrder.Id);
            newPurchaseOrder["hil_productdivision"] = data.ProductCategory;
            newPurchaseOrder["hil_salesoffice"] = data.SalesOffice;
            newPurchaseOrder["hil_franchise"] = data.Franchise;
            newPurchaseOrder["hil_warehouse"] = warehouse.ToEntityReference();
            newPurchaseOrder["ownerid"] = owner;
            newPurchaseOrder["hil_postatus"] = new OptionSetValue(1);
            newPurchaseOrder["hil_ordertype"] = new OptionSetValue(1);
            newPurchaseOrder["hil_brand"] = data.Brand;
            return service.Create(newPurchaseOrder);
        }

        private Entity GetFreshWarehouse(Guid franchise)
        {
            QueryExpression query = new QueryExpression("hil_inventorywarehouse");
            query.Criteria.AddCondition("hil_franchise", ConditionOperator.Equal, franchise);
            query.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 1); // Fresh

            EntityCollection warehouseColl = service.RetrieveMultiple(query);
            if (warehouseColl.Entities.Count > 0)
            {
                return warehouseColl.Entities[0];
            }
            return null;
        }

        internal void UpdateWorkOrder(Entity job)
        {
            string fetchworkOrder = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                <entity name='msdyn_workorder'>
                <attribute name='msdyn_name'/>
                <attribute name='msdyn_workorderid'/>
                <attribute name='msdyn_substatus'/>
                <order attribute='msdyn_name' descending='false'/>
                <filter type='and'>
                <condition attribute='msdyn_workorderid' operator='eq' value='{job.Id}'/>
                </filter>
                <link-entity name='hil_inventorypurchaseorderline' from='hil_workorder' to='msdyn_workorderid' link-type='inner' alias='ac'>
                <filter type='and'>
                <condition attribute='hil_partstatus' operator='not-in'>
                <value>3</value>
                <value>5</value>
                </condition>
                </filter>
                </link-entity>
                </entity>
                </fetch>";

            EntityCollection workOrderColl = service.RetrieveMultiple(new FetchExpression(fetchworkOrder)); 

            if(workOrderColl.Entities.Count == 0)
            {
                job["msdyn_substatus"] = workInitiated;
                service.Update(job); 
            }
        }
    }

    class JobData
    {
        public Entity WorkOrder { get; set; }
        public EntityReference SubStatus { get; set; }
        public EntityReference SalesOffice { get; set; }
        public EntityReference ProductCategory { get; set; }
        public EntityReference Franchise { get; set; }
        public OptionSetValue Brand { get; set; }
    }

    internal class InventoryRMALine
    {
        private IOrganizationService service;
        public InventoryRMALine(IOrganizationService _service)
        {
            this.service = _service;
        }

        private int GetTotalRMALines(Guid rmaId)
        {
            QueryExpression query = new QueryExpression("hil_inventoryrmaline");
            query.Criteria.AddCondition("hil_rma", ConditionOperator.Equal, rmaId);

            EntityCollection tempColl = service.RetrieveMultiple(query);

            return tempColl.Entities.Count + 1;
        }
        internal void UpdateLineNumber()
        {
            Entity inventoryRMALine = service.Retrieve("hil_inventoryrmaline", new Guid("53966dcb-cc3d-ef11-a316-6045bdcefd4c"), new ColumnSet("hil_name", "ownerid", "hil_rma"));
            EntityReference rmaRef = inventoryRMALine.GetAttributeValue<EntityReference>("hil_rma");

            Entity inventoryRMA = service.Retrieve("hil_inventoryrma", rmaRef.Id, new ColumnSet("hil_name"));
            string inventoryRMANumber = inventoryRMA.GetAttributeValue<string>("hil_name");
            Console.WriteLine(inventoryRMANumber);

            int seq = GetTotalRMALines(inventoryRMA.Id);

            inventoryRMALine["hil_name"] = inventoryRMANumber + "_" + seq;
            service.Update(inventoryRMALine);

        }
    }

    internal class InventoyWarehouse
    {
        private IOrganizationService service;
        public InventoyWarehouse(IOrganizationService _service)
        {
            this.service = _service;
        }
        internal void CreateInventoryWarehouse()
        {
            QueryExpression query = new QueryExpression("account");
            query.ColumnSet = new ColumnSet("hil_spareinventoryenabled", "accountnumber");

            EntityCollection tempColl = service.RetrieveMultiple(query);
            if (tempColl.Entities.Count > 0)
            {
                bool inventoryEnabled = tempColl.Entities[0].GetAttributeValue<Boolean>("hil_spareinventoryenabled");
                if (inventoryEnabled)
                {
                    string accountNumber = tempColl.Entities[0].GetAttributeValue<string>("accountnumber");
                    query = new QueryExpression("hil_inventorywarehouse");
                    query.Criteria.AddCondition("hil_name", ConditionOperator.Contains, accountNumber);

                    EntityCollection warehouseColl = service.RetrieveMultiple(query);
                    if (warehouseColl.Entities.Count == 0)
                    {
                        CreateWarehouse(accountNumber, "_Fresh");
                        CreateWarehouse(accountNumber, "_Defective");
                    }
                    else
                    {

                    }
                }
            }
        }

        private void CreateWarehouse(string accountNumber, string suffix)
        {
            Entity newWareHouse = new Entity("hil_inventorywarehouse");
            newWareHouse["hil_name"] = accountNumber + suffix;
            service.Create(newWareHouse);
        }
    }

    internal class PurchseOrder
    {
        private IOrganizationService service;
        public PurchseOrder(IOrganizationService _servce)
        {
            this.service = _servce;
        }

        private void notifyThroughEmail(Entity entity)
        {
            if (entity.Contains("hil_postatus"))
            {
                int poStatus = entity.GetAttributeValue<OptionSetValue>("hil_postatus").Value;
                if (poStatus == 2)
                {
                    // Send Email to approver 
                }
                else if (poStatus == 3 || poStatus == 6)
                {

                }
            }
        }


        public void preUpdateOfPurchaseOrder(Entity entity, IPluginExecutionContext context)
        {
            if (entity.Contains("hil_postatus"))
            {


            }

            Entity preImage = new Entity();
            if (context.PreEntityImages.Contains("hil_brand"))
            {
                preImage = context.PreEntityImages["hil_brand"];
            }
            int? preImagebrand = preImage.Contains("hil_brand") ? preImage.GetAttributeValue<OptionSetValue>("hil_brand").Value : null;
            int? brand = entity.Contains("hil_brand") ? entity.GetAttributeValue<OptionSetValue>("hil_brand").Value : null;

            if (brand != preImagebrand)
            {
                string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_inventorypurchaseorderline'>
                            <attribute name='hil_inventorypurchaseorderlineid' />
                            <attribute name='hil_name' />
                            <attribute name='createdon' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_ponumber' operator='eq' value='{entity.Id}' />
                              <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                          </entity>
                        </fetch>";
                EntityCollection entPurchaseorderline = service.RetrieveMultiple(new FetchExpression(fetchXML));

                if (entPurchaseorderline.Entities.Count > 0)
                {
                    throw new InvalidPluginExecutionException("Can't change the Brand! First deactivate all the line items to change.");
                }
            }

            Entity entPO = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_postatus"));
            int postatus = entPO.Contains("hil_postatus") ? entPO.GetAttributeValue<OptionSetValue>("hil_postatus").Value : 0;
            int[] IntPOstatus = { 2, 3, 4, 5, 6 };

            if (IntPOstatus.Contains(postatus))
            {
                throw new InvalidPluginExecutionException("You don't have authorization. Connect with system administrator");
            }
        }

        public void preUpdateOfPurchaseOrderLine(Entity entity, IPluginExecutionContext context)
        {
            Entity entPO = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_postatus"));
            Entity oridingalData = service.Retrieve(entPO.LogicalName, entPO.Id, new ColumnSet("hil_ponumber"));
            EntityReference purchaseOrderRef = oridingalData.GetAttributeValue<EntityReference>("hil_ponumber");
            Entity purchaseOrder = service.Retrieve(purchaseOrderRef.LogicalName, purchaseOrderRef.Id, new ColumnSet("hil_postatus", "ownerid"));

            int postatus = entPO.Contains("hil_postatus") ? entPO.GetAttributeValue<OptionSetValue>("hil_postatus").Value : 0;
            int[] IntPOstatus = [2, 3, 4, 5, 6]; // Submit for Approval, Approved, Posted , Cancelled, Rejected (in order) 

            if (IntPOstatus.Contains(postatus))
            {

                EntityReference owner = purchaseOrder.GetAttributeValue<EntityReference>("ownerid");

                if (owner.Id == context.InitiatingUserId && postatus == 2)
                {
                    if (entity.Attributes.Count > 1 || !entity.Contains("hil_quantity"))
                    {
                        throw new InvalidPluginExecutionException("You don't have authorization. Connect with system administrator");
                    }
                }
                else
                {
                    throw new InvalidPluginExecutionException("You don't have authorization. Connect with system administrator");
                }

            }
        }

    }

    internal class PostPurhaseOrderReceipt
    {
        private IOrganizationService service;
        public PostPurhaseOrderReceipt(IOrganizationService _service)
        {
            this.service = _service;
        }
        internal void PostOrderReceipt()
        {
            Entity orderReceipt = service.Retrieve("hil_inventorypurchaseorderreceipt", new Guid("c42475a7-574e-ef11-accd-6045bdad7f34"), new ColumnSet(true));
            List<Entity> receiptLines = GetReceiptLines(orderReceipt.Id);

            for (int i = 0; i < receiptLines.Count; i++)
            {
                int receiptQuantity = receiptLines[i].GetAttributeValue<int>("hil_receiptquantity");
                EntityReference orderLineRef = receiptLines[i].GetAttributeValue<EntityReference>("hil_purchaseorderline");
                Entity purchaseOrderLine = GetPurchaseOrderLine(orderLineRef.Id);
                int suppliedQuantity = purchaseOrderLine.Contains("hil_suppliedquantity") ? purchaseOrderLine.GetAttributeValue<int>("hil_suppliedquantity") : 0;
                purchaseOrderLine["hil_suppliedquantity"] = suppliedQuantity + receiptQuantity;
                service.Update(purchaseOrderLine);
            }

        }

        private Entity GetPurchaseOrderLine(Guid orderLineId)
        {
            return service.Retrieve("hil_inventorypurchaseorderline", orderLineId, new ColumnSet(true));
        }

        private List<Entity> GetReceiptLines(Guid orderReceiptId)
        {
            QueryExpression expression = new QueryExpression("hil_inventorypurchaseorderreceiptline");
            expression.ColumnSet = new ColumnSet("hil_receiptquantity", "hil_purchaseorderline");
            expression.Criteria.AddCondition("hil_receiptnumber", ConditionOperator.Equal, orderReceiptId);
            return service.RetrieveMultiple(expression).Entities.ToList<Entity>();
        }
    }
    
}
