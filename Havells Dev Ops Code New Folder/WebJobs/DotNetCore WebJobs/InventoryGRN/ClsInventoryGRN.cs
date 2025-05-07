using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Net;
using System.Runtime.InteropServices.JavaScript;
using System.Text;

namespace InventoryGRN
{
    public class ClsInventoryGRN
    {
        private readonly ServiceClient _service;
        public ClsInventoryGRN(ServiceClient service)
        {
            _service = service;
        }
        //public void CreateGRNLine(string _startDate, string _endDate)
        //{
        //    IntegrationConfiguration integrationConfiguration = GetIntegrationConfiguration("CRM_SAPtoCRM_Invoice");
        //    integrationConfiguration.url = integrationConfiguration.url + $"&IM_FROM_DT={_startDate}&IM_TO_DT={_endDate}";
        //    //   integrationConfiguration.url = integrationConfiguration.url.Replace("middleware", "p90ci") + $"&IM_FROM_DT={_startDate}&IM_TO_DT={_endDate}";
        //    InvoiceResonse resp = JsonConvert.DeserializeObject<InvoiceResonse>(CallAPI(integrationConfiguration, null, "GET"));

        //    if (resp.LT_INVOICE.Count > 0 && resp.LT_INVOICE[0] != null)
        //    {
        //        QueryExpression query;
        //        int recNo = 1;
        //        foreach (LTINVOICE lTINVOICE in resp.LT_INVOICE)
        //        {
        //            try
        //            {
        //                Console.WriteLine($"Record# {recNo++}/{resp.LT_INVOICE.Count} Invoice# {lTINVOICE.VBELN} SO# {lTINVOICE.AUBEL}");
        //                if (lTINVOICE != null)
        //                {
        //                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                        <entity name='hil_inventorypurchaseorder'>
        //                        <attribute name='hil_inventorypurchaseorderid' />
        //                        <order attribute='hil_name' descending='false' />
        //                        <filter type='and'>
        //                            <condition attribute='hil_salesordernumber' operator='eq' value='{lTINVOICE.AUBEL.Trim()}' />
        //                        </filter>
        //                        </entity>
        //                        </fetch>";
        //                    EntityCollection entColPO = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                    if (entColPO.Entities.Count > 0)
        //                    {
        //                        string franchiseCode = lTINVOICE.KUNRG.Substring(1);

        //                        EntityReference product = GetProduct(lTINVOICE.MATNR);
        //                        EntityReference franchise = GetFranchise(lTINVOICE.KUNRG);

        //                        if (product != null && franchise != null)
        //                        {
        //                            query = new QueryExpression("hil_inventorypurchaseorder");
        //                            query.ColumnSet = new ColumnSet("ownerid", "hil_jobid");
        //                            query.Criteria.AddCondition("hil_salesordernumber", ConditionOperator.Equal, lTINVOICE.AUBEL);
        //                            EntityCollection purchaseOrderCollection = _service.RetrieveMultiple(query);

        //                            query = new QueryExpression("hil_inventorywarehouse");
        //                            query.Criteria.AddCondition("hil_franchise", ConditionOperator.Equal, franchise.Id);
        //                            query.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 1); //fresh 
        //                            EntityCollection wareHouseCollection = _service.RetrieveMultiple(query);

        //                            if (purchaseOrderCollection.Entities.Count > 0 && wareHouseCollection.Entities.Count > 0)
        //                            {
        //                                query = new QueryExpression("hil_inventorysparebills");
        //                                query.Criteria.AddCondition("hil_salesordernumber", ConditionOperator.Equal, lTINVOICE.AUBEL);
        //                                query.Criteria.AddCondition("hil_billnumber", ConditionOperator.Equal, lTINVOICE.VBELN);
        //                                query.Criteria.AddCondition("hil_billdate", ConditionOperator.Equal, lTINVOICE.FKDAT);
        //                                query.Criteria.AddCondition("hil_franchise", ConditionOperator.Equal, franchise.Id);
        //                                query.Criteria.AddCondition("hil_partcode", ConditionOperator.Equal, product.Id);
        //                                EntityCollection doesBillExist = _service.RetrieveMultiple(query);
        //                                if (doesBillExist.Entities.Count == 0)
        //                                {
        //                                    Entity newInventorySpareBills = new Entity("hil_inventorysparebills");
        //                                    newInventorySpareBills["hil_salesordernumber"] = lTINVOICE.AUBEL;
        //                                    newInventorySpareBills["hil_billnumber"] = lTINVOICE.VBELN;
        //                                    newInventorySpareBills["hil_billdate"] = lTINVOICE.FKDAT;
        //                                    if (!string.IsNullOrEmpty(lTINVOICE.FKIMG))
        //                                    {
        //                                        newInventorySpareBills["hil_quantity"] = Convert.ToInt32(Convert.ToDecimal(lTINVOICE.FKIMG));//Quantity
        //                                    }
        //                                    newInventorySpareBills["hil_franchise"] = franchise;
        //                                    newInventorySpareBills["hil_partcode"] = product;

        //                                    if (purchaseOrderCollection.Entities.Count > 0)
        //                                    {
        //                                        newInventorySpareBills["hil_ordernumber"] = purchaseOrderCollection.Entities[0].ToEntityReference();
        //                                    }
        //                                    _service.Create(newInventorySpareBills);
        //                                }

        //                                EntityReference workOrderRef = purchaseOrderCollection.Entities[0].Contains("hil_jobid") ? purchaseOrderCollection.Entities[0].GetAttributeValue<EntityReference>("hil_jobid") : null;

        //                                EntityReference poOwner = purchaseOrderCollection.Entities[0].GetAttributeValue<EntityReference>("ownerid");

        //                                string _fetchxml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                                  <entity name='hil_inventorypurchaseorderreceiptline'>
        //                                    <attribute name='hil_inventorypurchaseorderreceiptlineid' />
        //                                    <attribute name='hil_name' />
        //                                    <attribute name='createdon' />
        //                                    <order attribute='hil_name' descending='false' />
        //                                    <filter type='and'>
        //                                      <condition attribute='hil_ordernumber' operator='eq' value='{purchaseOrderCollection.Entities[0].Id}' />
        //                                      <condition attribute='hil_partcode' operator='eq' value='{product.Id}' />
        //                                    </filter>
        //                                    <link-entity name='hil_inventorypurchaseorderreceipt' from='hil_inventorypurchaseorderreceiptid' to='hil_receiptnumber' link-type='inner' alias='ac'>
        //                                      <filter type='and'>
        //                                        <condition attribute='hil_invoicenumber' operator='eq' value='{lTINVOICE.VBELN.Trim()}' />
        //                                        <condition attribute='hil_sapsonumber' operator='eq' value='{lTINVOICE.AUBEL.Trim()}' />
        //                                      </filter>
        //                                    </link-entity>
        //                                  </entity>
        //                                </fetch>";
        //                                EntityCollection _entCol = _service.RetrieveMultiple(new FetchExpression(_fetchxml));
        //                                if (_entCol.Entities.Count > 0)
        //                                {
        //                                    continue;
        //                                }
        //                                else
        //                                {
        //                                    Guid purchaseOrderReceipt = GetInventoryPurchaseRecept(lTINVOICE.VBELN, lTINVOICE.AUBEL, franchise, purchaseOrderCollection.Entities[0].ToEntityReference(), wareHouseCollection.Entities[0].ToEntityReference(), poOwner.Id, workOrderRef);
        //                                    int billedQuantity = Convert.ToInt32(Convert.ToDecimal(lTINVOICE.FKIMG));//Quantity
        //                                    Guid purchaseOrderReceiptLine = GetPurchaseOrderReceiptLine(lTINVOICE.VBELN, lTINVOICE.AUBEL, purchaseOrderReceipt, product, purchaseOrderCollection.Entities[0].ToEntityReference(), wareHouseCollection.Entities[0].ToEntityReference(), poOwner.Id, workOrderRef, billedQuantity);
        //                                }
        //                            }
        //                        }
        //                    }
        //                    else
        //                    {
        //                        Console.WriteLine($"SO# {lTINVOICE.AUBEL.Trim()} doesn't exist in CRM.");
        //                    }
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("error " + ex.Message);
        //            }
        //        }
        //    }
        //}
        //public void CreateGRNLineBySONumber(string SONumber)
        //{
        //    IntegrationConfiguration integrationConfiguration = GetIntegrationConfiguration("CRM_SAPtoCRM_Invoice");
        //    integrationConfiguration.url = integrationConfiguration.url + $"&IM_VBELN={SONumber}";
        //    // integrationConfiguration.url = integrationConfiguration.url.Replace("middleware", "p90ci") + $"&IM_VBELN={SONumber}";
        //    InvoiceResonse resp = JsonConvert.DeserializeObject<InvoiceResonse>(CallAPI(integrationConfiguration, null, "GET"));

        //    if (resp.LT_INVOICE.Count > 0 && resp.LT_INVOICE[0] != null)
        //    {
        //        QueryExpression query;
        //        int recNo = 1;
        //        foreach (LTINVOICE lTINVOICE in resp.LT_INVOICE)
        //        {
        //            try
        //            {
        //                Console.WriteLine($"Record# {recNo++}/{resp.LT_INVOICE.Count} Invoice# {lTINVOICE.VBELN} SO# {lTINVOICE.AUBEL}");
        //                if (lTINVOICE != null)
        //                {
        //                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                        <entity name='hil_inventorypurchaseorder'>
        //                        <attribute name='hil_inventorypurchaseorderid' />
        //                        <order attribute='hil_name' descending='false' />
        //                        <filter type='and'>
        //                            <condition attribute='hil_salesordernumber' operator='eq' value='{lTINVOICE.AUBEL.Trim()}' />
        //                        </filter>
        //                        </entity>
        //                        </fetch>";
        //                    EntityCollection entColPO = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                    if (entColPO.Entities.Count > 0)
        //                    {
        //                        string franchiseCode = lTINVOICE.KUNRG.Substring(1);

        //                        EntityReference product = GetProduct(lTINVOICE.MATNR);
        //                        EntityReference franchise = GetFranchise(lTINVOICE.KUNRG);

        //                        if (product != null && franchise != null)
        //                        {
        //                            query = new QueryExpression("hil_inventorypurchaseorder");
        //                            query.ColumnSet = new ColumnSet("ownerid", "hil_jobid");
        //                            query.Criteria.AddCondition("hil_salesordernumber", ConditionOperator.Equal, lTINVOICE.AUBEL);
        //                            EntityCollection purchaseOrderCollection = _service.RetrieveMultiple(query);
        //                            query = new QueryExpression("hil_inventorywarehouse");
        //                            query.Criteria.AddCondition("hil_franchise", ConditionOperator.Equal, franchise.Id);
        //                            query.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 1); //fresh 
        //                            EntityCollection wareHouseCollection = _service.RetrieveMultiple(query);

        //                            if (purchaseOrderCollection.Entities.Count > 0 && wareHouseCollection.Entities.Count > 0)
        //                            {
        //                                query = new QueryExpression("hil_inventorysparebills");
        //                                query.Criteria.AddCondition("hil_salesordernumber", ConditionOperator.Equal, lTINVOICE.AUBEL);
        //                                query.Criteria.AddCondition("hil_billnumber", ConditionOperator.Equal, lTINVOICE.VBELN);
        //                                query.Criteria.AddCondition("hil_billdate", ConditionOperator.Equal, lTINVOICE.FKDAT);
        //                                query.Criteria.AddCondition("hil_franchise", ConditionOperator.Equal, franchise.Id);
        //                                query.Criteria.AddCondition("hil_partcode", ConditionOperator.Equal, product.Id);
        //                                EntityCollection doesBillExist = _service.RetrieveMultiple(query);
        //                                if (doesBillExist.Entities.Count == 0)
        //                                {
        //                                    Entity newInventorySpareBills = new Entity("hil_inventorysparebills");
        //                                    newInventorySpareBills["hil_salesordernumber"] = lTINVOICE.AUBEL;
        //                                    newInventorySpareBills["hil_billnumber"] = lTINVOICE.VBELN;
        //                                    newInventorySpareBills["hil_billdate"] = lTINVOICE.FKDAT;
        //                                    if (!string.IsNullOrEmpty(lTINVOICE.FKIMG))
        //                                    {
        //                                        newInventorySpareBills["hil_quantity"] = Convert.ToInt32(Convert.ToDecimal(lTINVOICE.FKIMG));//Quantity
        //                                    }
        //                                    newInventorySpareBills["hil_franchise"] = franchise;
        //                                    newInventorySpareBills["hil_partcode"] = product;

        //                                    if (purchaseOrderCollection.Entities.Count > 0)
        //                                    {
        //                                        newInventorySpareBills["hil_ordernumber"] = purchaseOrderCollection.Entities[0].ToEntityReference();
        //                                    }
        //                                    _service.Create(newInventorySpareBills);
        //                                }

        //                                EntityReference workOrderRef = purchaseOrderCollection.Entities[0].Contains("hil_jobid") ? purchaseOrderCollection.Entities[0].GetAttributeValue<EntityReference>("hil_jobid") : null;

        //                                EntityReference poOwner = purchaseOrderCollection.Entities[0].GetAttributeValue<EntityReference>("ownerid");

        //                                string _fetchxml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                                  <entity name='hil_inventorypurchaseorderreceiptline'>
        //                                    <attribute name='hil_inventorypurchaseorderreceiptlineid' />
        //                                    <attribute name='hil_name' />
        //                                    <attribute name='createdon' />
        //                                    <order attribute='hil_name' descending='false' />
        //                                    <filter type='and'>
        //                                      <condition attribute='hil_ordernumber' operator='eq' value='{purchaseOrderCollection.Entities[0].Id}' />
        //                                      <condition attribute='hil_partcode' operator='eq' value='{product.Id}' />
        //                                    </filter>
        //                                    <link-entity name='hil_inventorypurchaseorderreceipt' from='hil_inventorypurchaseorderreceiptid' to='hil_receiptnumber' link-type='inner' alias='ac'>
        //                                      <filter type='and'>
        //                                        <condition attribute='hil_invoicenumber' operator='eq' value='{lTINVOICE.VBELN.Trim()}' />
        //                                        <condition attribute='hil_sapsonumber' operator='eq' value='{lTINVOICE.AUBEL.Trim()}' />
        //                                      </filter>
        //                                    </link-entity>
        //                                  </entity>
        //                                </fetch>";
        //                                EntityCollection _entCol = _service.RetrieveMultiple(new FetchExpression(_fetchxml));
        //                                if (_entCol.Entities.Count > 0)
        //                                {
        //                                    continue;
        //                                }
        //                                else
        //                                {
        //                                    Guid purchaseOrderReceipt = GetInventoryPurchaseRecept(lTINVOICE.VBELN, lTINVOICE.AUBEL, franchise, purchaseOrderCollection.Entities[0].ToEntityReference(), wareHouseCollection.Entities[0].ToEntityReference(), poOwner.Id, workOrderRef);
        //                                    int billedQuantity = Convert.ToInt32(Convert.ToDecimal(lTINVOICE.FKIMG));//Quantity
        //                                    Guid purchaseOrderReceiptLine = GetPurchaseOrderReceiptLine(lTINVOICE.VBELN, lTINVOICE.AUBEL, purchaseOrderReceipt, product, purchaseOrderCollection.Entities[0].ToEntityReference(), wareHouseCollection.Entities[0].ToEntityReference(), poOwner.Id, workOrderRef, billedQuantity);
        //                                }
        //                            }
        //                        }
        //                    }
        //                    else
        //                    {
        //                        Console.WriteLine($"SO# {lTINVOICE.AUBEL.Trim()} doesn't exist in CRM.");
        //                    }
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("error " + ex.Message);
        //            }
        //        }
        //    }
        //}

        public void CreateGRNLineBySONumberNew(string SONumber)
        {
            var Json = JsonConvert.SerializeObject(new { IM_VBELN = SONumber });
            IntegrationConfiguration integrationConfiguration = GetIntegrationConfiguration("CRM_SAPtoCRM_InvoiceNew");
            integrationConfiguration.url = integrationConfiguration.url.Replace("middleware", "p90ci");
            InvoiceResonse resp = JsonConvert.DeserializeObject<InvoiceResonse>(CallAPI(integrationConfiguration, Json, "POST"));
            if (resp.LT_INVOICE.Count > 0 && resp.LT_INVOICE[0] != null)
            {
                QueryExpression query;
                int recNo = 1;
                foreach (LTINVOICE lTINVOICE in resp.LT_INVOICE)
                {
                    try
                    {
                        Console.WriteLine($"Record# {recNo++}/{resp.LT_INVOICE.Count} Invoice# {lTINVOICE.BIL_NO} SO# {lTINVOICE.ORD_NO}");
                        if (lTINVOICE != null)
                        {
                            if (string.IsNullOrWhiteSpace(lTINVOICE.ORD_NO))
                            {
                                Console.WriteLine(JsonConvert.SerializeObject(resp));
                                Console.WriteLine($"{resp.LT_ELOG[0].ERROR} for SO# {SONumber}");
                                continue;
                            }
                            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_inventorypurchaseorder'>
                                <attribute name='hil_inventorypurchaseorderid' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                    <condition attribute='hil_salesordernumber' operator='eq' value='{lTINVOICE.ORD_NO.Trim()}' />
                                </filter>
                                </entity>
                                </fetch>";
                            EntityCollection entColPO = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entColPO.Entities.Count > 0)
                            {
                                EntityReference product = GetProduct(lTINVOICE.MATNR);
                                EntityReference franchise = GetFranchise(lTINVOICE.KUNNR);
                                string franchiseCode = lTINVOICE.KUNNR.Substring(1);
                                int billedQuantity = !string.IsNullOrWhiteSpace(lTINVOICE.FKIMG) ? Convert.ToInt32(Convert.ToDecimal(lTINVOICE.FKIMG.Trim())) : 0;

                                if (product != null && franchise != null)
                                {
                                    query = new QueryExpression("hil_inventorypurchaseorder");
                                    query.ColumnSet = new ColumnSet("ownerid", "hil_jobid");
                                    query.Criteria.AddCondition("hil_salesordernumber", ConditionOperator.Equal, lTINVOICE.ORD_NO);
                                    EntityCollection purchaseOrderCollection = _service.RetrieveMultiple(query);

                                    query = new QueryExpression("hil_inventorywarehouse");
                                    query.Criteria.AddCondition("hil_franchise", ConditionOperator.Equal, franchise.Id);
                                    query.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 1); //fresh 
                                    EntityCollection wareHouseCollection = _service.RetrieveMultiple(query);

                                    if (purchaseOrderCollection.Entities.Count > 0 && wareHouseCollection.Entities.Count > 0)
                                    {
                                        if (Convert.ToInt32(lTINVOICE.FKDAT) == 0)//Part not dispatched
                                        {
                                            Console.WriteLine($"Part : {lTINVOICE.MATNR} not dispatched.");
                                            continue;
                                        }
                                        if (!string.IsNullOrWhiteSpace(lTINVOICE.REJ_RESN))//Rejection Reason
                                        {
                                            Console.WriteLine($"Part code {lTINVOICE.MATNR} Rejection Reason : {lTINVOICE.REJ_RESN}.");
                                            continue;
                                        }

                                        string dateAsString = lTINVOICE.FKDAT.ToString();
                                        string year = dateAsString.Substring(0, 4);
                                        string day = dateAsString.Substring(6, 2);
                                        string month = dateAsString.Substring(4, 2);
                                        string formattedFKDAT = $"{year}-{month}-{day}";

                                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                  <entity name='hil_inventorysparebills'>
                                                    <attribute name='hil_inventorysparebillsid' />    
                                                    <attribute name='hil_salesordernumber' />
                                                    <attribute name='hil_billnumber' />
                                                    <attribute name='hil_billdate' />
                                                    <attribute name='hil_ordernumber' />
                                                    <attribute name='hil_partcode' />
                                                    <attribute name='hil_franchise' />
                                                    <attribute name='hil_quantity' />
                                                    <order attribute='hil_billnumber' descending='false' />
                                                    <filter type='and'>
                                                      <condition attribute='hil_salesordernumber' operator='eq' value='{lTINVOICE.ORD_NO}' />
                                                      <condition attribute='hil_billnumber' operator='eq' value='{lTINVOICE.BIL_NO}' />
                                                      <condition attribute='hil_billdate' operator='on' value='{formattedFKDAT}' />
                                                      <condition attribute='hil_franchise' operator='eq' value='{franchise.Id}' />
                                                      <condition attribute='hil_partcode' operator='eq' value='{product.Id}' />
                                                    </filter>
                                                  </entity>
                                                </fetch>";
                                        EntityCollection doesBillExist = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                        if (doesBillExist.Entities.Count == 0)
                                        {
                                            Entity newInventorySpareBills = new Entity("hil_inventorysparebills");
                                            newInventorySpareBills["hil_salesordernumber"] = lTINVOICE.ORD_NO;
                                            newInventorySpareBills["hil_billnumber"] = lTINVOICE.BIL_NO.ToString();
                                            newInventorySpareBills["hil_billdate"] = new DateTime(Convert.ToInt32(year), Convert.ToInt32(month), Convert.ToInt32(day));
                                            newInventorySpareBills["hil_franchise"] = franchise;
                                            newInventorySpareBills["hil_partcode"] = product;
                                            newInventorySpareBills["hil_ordernumber"] = purchaseOrderCollection.Entities[0].ToEntityReference();
                                            newInventorySpareBills["hil_quantity"] = billedQuantity;//Quantity
                                            _service.Create(newInventorySpareBills);
                                        }

                                        EntityReference workOrderRef = purchaseOrderCollection.Entities[0].Contains("hil_jobid") ? purchaseOrderCollection.Entities[0].GetAttributeValue<EntityReference>("hil_jobid") : null;
                                        EntityReference poOwner = purchaseOrderCollection.Entities[0].GetAttributeValue<EntityReference>("ownerid");

                                        string _fetchxml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='hil_inventorypurchaseorderreceiptline'>
                                            <attribute name='hil_inventorypurchaseorderreceiptlineid' />
                                            <attribute name='hil_name' />
                                            <attribute name='createdon' />
                                            <order attribute='hil_name' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='hil_ordernumber' operator='eq' value='{purchaseOrderCollection.Entities[0].Id}' />
                                              <condition attribute='hil_partcode' operator='eq' value='{product.Id}' />
                                            </filter>
                                            <link-entity name='hil_inventorypurchaseorderreceipt' from='hil_inventorypurchaseorderreceiptid' to='hil_receiptnumber' link-type='inner' alias='ac'>
                                              <filter type='and'>
                                                <condition attribute='hil_invoicenumber' operator='eq' value='{lTINVOICE.BIL_NO.ToString().Trim()}' />
                                                <condition attribute='hil_sapsonumber' operator='eq' value='{lTINVOICE.ORD_NO.Trim()}' />
                                              </filter>
                                            </link-entity>
                                          </entity>
                                        </fetch>";
                                        EntityCollection _entCol = _service.RetrieveMultiple(new FetchExpression(_fetchxml));
                                        if (_entCol.Entities.Count > 0)
                                        {
                                            Console.WriteLine($"Purchase order receipt line : {_entCol.Entities[0].Id} already exist.");
                                            continue;
                                        }
                                        else
                                        {
                                            PurchaseOrderReceipt ObjPurchaseOrderReceipt = new PurchaseOrderReceipt
                                            {
                                                SapInvoiceNumber = lTINVOICE.BIL_NO.ToString().Trim(),
                                                SapOrderNumber = lTINVOICE.ORD_NO.Trim(),
                                                Franchise = franchise,
                                                PurchaseOrder = purchaseOrderCollection.Entities[0].ToEntityReference(),
                                                FreshWarehouse = wareHouseCollection.Entities[0].ToEntityReference(),
                                                OwnerId = poOwner.Id,
                                                WorkOrderRef = workOrderRef
                                            };
                                            Guid purchaseOrderReceiptId = GetInventoryPurchaseRecept(ObjPurchaseOrderReceipt);
                                            PurchaseOrderReceiptLine ObjPurchaseOrderReceiptLine = new PurchaseOrderReceiptLine
                                            {
                                                PurchaseOrderReceiptDetails = ObjPurchaseOrderReceipt,
                                                PurchaseOrderReceiptId = purchaseOrderReceiptId,
                                                Product = product,
                                                BilledQuantity = billedQuantity,
                                                SAPOrderItem = lTINVOICE.ORD_ITEM.Trim()
                                            };
                                            GetPurchaseOrderReceiptLine(ObjPurchaseOrderReceiptLine);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine($"SO# {lTINVOICE.ORD_NO.Trim()} doesn't exist in CRM.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error : {ex.Message}");
                        continue;
                    }
                }
            }
        }
        public void CreateGRNLineForSONumberGeneratedInLast_Hr(string Hr)
        {
            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                  <entity name='hil_inventorypurchaseorder'>
                                    <attribute name='hil_name' />
                                    <attribute name='hil_salesordernumber' />
                                    <attribute name='hil_jobid' />
                                    <attribute name='hil_franchise' />
                                    <attribute name='hil_inventorypurchaseorderid' />
                                    <attribute name='hil_warehouse' />
                                    <order attribute='createdon' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='statecode' operator='eq' value='0' />
                                      <condition attribute='hil_salesordernumber' operator='not-null' />
                                      <condition attribute='hil_postatus' operator='eq' value='4' />
                                      <condition attribute='createdon' operator='last-x-hours' value='{Hr}' />
                                    </filter>
                                    <link-entity name='account' from='accountid' to='hil_franchise' visible='false' link-type='outer' alias='ac'>
                                      <attribute name='accountnumber' />
                                    </link-entity>
                                  </entity>
                                </fetch>";
            EntityCollection entColPO = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entColPO.Entities.Count > 0)
            {
                int rowCount = 0;
                foreach (var entPO in entColPO.Entities)
                {
                    Console.WriteLine($"{rowCount++}/{entColPO.Entities.Count}|{entPO.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString()}");
                    // CreateGRNLineBySONumber(entPO.GetAttributeValue<string>("hil_salesordernumber"));
                    try
                    {
                        CreateGRNLineBySONumberNew(entPO.GetAttributeValue<string>("hil_salesordernumber"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error : {ex.Message}");
                        continue;
                    }
                }
            }
        }

        #region Supporting Methods
        private IntegrationConfiguration GetIntegrationConfiguration(string APIName)
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
        private string CallAPI(IntegrationConfiguration integrationConfiguration, string Json, String method)
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
            Stream dataStream1 = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream1);
            return reader.ReadToEnd();
        }
        private EntityReference? GetFranchise(string franchiseCode)
        {
            franchiseCode = franchiseCode[..1] == "F" ? franchiseCode[1..] : franchiseCode;
            QueryExpression query = new QueryExpression("account");
            query.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, franchiseCode);
            EntityCollection tempCollection = _service.RetrieveMultiple(query);
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
            EntityCollection tempCollection = _service.RetrieveMultiple(query);
            if (tempCollection.Entities.Count > 0)
            {
                return tempCollection.Entities[0].ToEntityReference();
            }
            return null;
        }
        private Guid GetInventoryPurchaseRecept(PurchaseOrderReceipt purchaseOrderReceipt)
        {
            QueryExpression query = new QueryExpression("hil_inventorypurchaseorderreceipt");
            query.Criteria.AddCondition("hil_invoicenumber", ConditionOperator.Equal, purchaseOrderReceipt.SapInvoiceNumber);
            query.Criteria.AddCondition("hil_sapsonumber", ConditionOperator.Equal, purchaseOrderReceipt.SapOrderNumber);
            query.Criteria.AddCondition("hil_receiptstatus", ConditionOperator.Equal, 1); // draft

            EntityCollection receiptCollection = _service.RetrieveMultiple(query);
            if (receiptCollection.Entities.Count > 0)
            {
                return receiptCollection.Entities[0].Id;
            }
            Entity newPurchaseReceipt = new Entity("hil_inventorypurchaseorderreceipt");
            newPurchaseReceipt["hil_franchise"] = purchaseOrderReceipt.Franchise;
            newPurchaseReceipt["hil_invoicenumber"] = purchaseOrderReceipt.SapInvoiceNumber;
            newPurchaseReceipt["hil_sapsonumber"] = purchaseOrderReceipt.SapOrderNumber;
            newPurchaseReceipt["hil_ordernumber"] = purchaseOrderReceipt.PurchaseOrder;
            newPurchaseReceipt["hil_warehouse"] = purchaseOrderReceipt.FreshWarehouse;
            newPurchaseReceipt["hil_receiptstatus"] = new OptionSetValue(1);
            newPurchaseReceipt["ownerid"] = new EntityReference("systemuser", purchaseOrderReceipt.OwnerId);
            if (purchaseOrderReceipt.WorkOrderRef != null) newPurchaseReceipt["hil_jobid"] = purchaseOrderReceipt.WorkOrderRef;
            return _service.Create(newPurchaseReceipt);
        }
        private Entity GetOrderLine(EntityReference product, EntityReference purhaseOrder)
        {
            QueryExpression query = new QueryExpression("hil_inventorypurchaseorderline");
            query.Criteria.AddCondition("hil_partcode", ConditionOperator.Equal, product.Id);
            query.Criteria.AddCondition("hil_ponumber", ConditionOperator.Equal, purhaseOrder.Id);

            EntityCollection orderLineColl = _service.RetrieveMultiple(query);
            return orderLineColl.Entities[0];
        }
        private Guid GetPurchaseOrderReceiptLine(PurchaseOrderReceiptLine purchaseOrderReceiptLine)
        {
            QueryExpression query = new QueryExpression("hil_inventorypurchaseorderreceiptline");
            query.Criteria.AddCondition("hil_receiptnumber", ConditionOperator.Equal, purchaseOrderReceiptLine.PurchaseOrderReceiptId);
            query.Criteria.AddCondition("hil_partcode", ConditionOperator.Equal, purchaseOrderReceiptLine.Product.Id);

            EntityCollection linesCollection = _service.RetrieveMultiple(query);
            if (linesCollection.Entities.Count > 0)
            {
                return linesCollection.Entities[0].Id;
            }

            Entity newPurchaseLine = new Entity("hil_inventorypurchaseorderreceiptline");
            newPurchaseLine["hil_receiptnumber"] = new EntityReference("hil_inventorypurchaseorderreceipt", purchaseOrderReceiptLine.PurchaseOrderReceiptId);
            newPurchaseLine["hil_sapbillnumber"] = purchaseOrderReceiptLine.PurchaseOrderReceiptDetails.SapInvoiceNumber;
            newPurchaseLine["hil_sapsonumber"] = purchaseOrderReceiptLine.PurchaseOrderReceiptDetails.SapOrderNumber;
            newPurchaseLine["hil_saporderitem"] = purchaseOrderReceiptLine.SAPOrderItem;
            newPurchaseLine["hil_partcode"] = purchaseOrderReceiptLine.Product;
            newPurchaseLine["hil_ordernumber"] = purchaseOrderReceiptLine.PurchaseOrderReceiptDetails.PurchaseOrder;
            newPurchaseLine["hil_warehouse"] = purchaseOrderReceiptLine.PurchaseOrderReceiptDetails.FreshWarehouse;
            newPurchaseLine["hil_purchaseorderline"] = GetOrderLine(purchaseOrderReceiptLine.Product, purchaseOrderReceiptLine.PurchaseOrderReceiptDetails.PurchaseOrder).ToEntityReference();
            newPurchaseLine["ownerid"] = new EntityReference("systemuser", purchaseOrderReceiptLine.PurchaseOrderReceiptDetails.OwnerId);
            if (purchaseOrderReceiptLine.PurchaseOrderReceiptDetails.WorkOrderRef != null) newPurchaseLine["hil_jobid"] = purchaseOrderReceiptLine.PurchaseOrderReceiptDetails.WorkOrderRef;
            newPurchaseLine["hil_billedquantity"] = purchaseOrderReceiptLine.BilledQuantity;
            return _service.Create(newPurchaseLine);
        }
        #endregion
    }
    public class IntegrationConfiguration
    {
        public string url { get; set; }
        public string userName { get; set; }
        public string password { get; set; }

    }
    //public class InvoiceResonse
    //{
    //    public List<LTINVOICE> LT_INVOICE { get; set; }
    //}
    //public class LTINVOICE
    //{
    //    public string VBELN { get; set; }
    //    public string POSNR { get; set; }
    //    public decimal ECGST_PERC { get; set; }
    //    public string FKART { get; set; }
    //    public DateTime FKDAT { get; set; }
    //    public string AUBEL { get; set; }
    //    public string NETWR { get; set; }
    //    public string MWSBK { get; set; }
    //    public string TOT_DISC { get; set; }
    //    public string KUNRG { get; set; }
    //    public string WERKS { get; set; }
    //    public string BUTXT { get; set; }
    //    public string BEZEI { get; set; }
    //    public string MATNR { get; set; }
    //    public string FKIMG { get; set; }
    //    public string MEINS { get; set; }
    //    public string CHARG { get; set; }
    //    public string MRP { get; set; }
    //    public string DLP { get; set; }
    //    public string PUR_RATE { get; set; }
    //    public string LLAMOUNT { get; set; }
    //    public string LLDISCOUNT { get; set; }
    //    public string LLTAXAMT { get; set; }
    //    public string LLTOTAMT { get; set; }
    //    public string BT_MANUF_DATE { get; set; }
    //    public string BT_EXP_DATE { get; set; }
    //    public string ROUNDOFF { get; set; }
    //    public string FKSTO { get; set; }
    //    public string TAXABLE_AMT { get; set; }
    //    public string CGST_PERC { get; set; }
    //    public string SGST_PERC { get; set; }
    //    public decimal UTGST_PERC { get; set; }
    //    public string IGST_PERC { get; set; }
    //    public string CESS_PERC { get; set; }
    //    public string CGST_AMT { get; set; }
    //    public string SGST_AMT { get; set; }
    //    public decimal UTGST_AMT { get; set; }
    //    public string IGST_AMT { get; set; }
    //    public int STEUC { get; set; }
    //    public string REL_PARTY { get; set; }
    //    public string INV_TYPE { get; set; }
    //}






    public class LTELOG
    {
        public string KEY1 { get; set; }
        public string KEY2 { get; set; }
        public string KEY3 { get; set; }
        public string ERROR { get; set; }
    }

    public class LTINVOICE
    {
        public string ORD_NO { get; set; }
        public object BIL_NO { get; set; }
        public dynamic FKDAT { get; set; }
        public string MATNR { get; set; }
        public string FKIMG { get; set; }
        public string KUNNR { get; set; }
        public string ORD_ITEM { get; set; }
        public string REJ_RESN { get; set; }

    }
    public class InvoiceResonse
    {
        public List<LTINVOICE> LT_INVOICE { get; set; }
        public List<LTELOG> LT_ELOG { get; set; }
    }

    public class PurchaseOrderReceipt
    {
        public Guid OwnerId { get; set; }
        public string SapInvoiceNumber { get; set; }
        public string SapOrderNumber { get; set; }
        public EntityReference PurchaseOrder { get; set; }
        public EntityReference WorkOrderRef { get; set; }
        public EntityReference Franchise { get; set; }
        public EntityReference FreshWarehouse { get; set; }

    }

    public class PurchaseOrderReceiptLine
    {
        public PurchaseOrderReceipt PurchaseOrderReceiptDetails { get; set; }
        public Guid PurchaseOrderReceiptId { get; set; }
        public EntityReference Product { get; set; }
        public int BilledQuantity { get; set; }
        public string SAPOrderItem { get; set; }
    }

}
