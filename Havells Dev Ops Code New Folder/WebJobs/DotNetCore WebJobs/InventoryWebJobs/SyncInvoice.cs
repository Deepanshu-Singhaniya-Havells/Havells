using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Rest;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Collections.Generic;
using System;

namespace InventoryWebJobs
{

    internal class SyncInvoice
    {
        private readonly EntityReference workInitiated = new EntityReference("msdyn_workordersubstatus", new Guid("2b27fa6c-fa0f-e911-a94e-000d3af060a1"));
        private readonly ServiceClient service;

        public SyncInvoice(ServiceClient _service)
        {
            this.service = _service;
        }

        private class InvoiceResonse
        {
            public List<LTINVOICE> LT_INVOICE { get; set; }
        }

        private class LTINVOICE
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

        private void UpdateWorkOrder(EntityReference workOrderRef)
        {
            Entity woObj = service.Retrieve(workOrderRef.LogicalName, workOrderRef.Id, new ColumnSet("msdyn_substatus"));
            EntityReference _jobSubStatus = woObj.GetAttributeValue<EntityReference>("msdyn_substatus");

            if (_jobSubStatus.Id == new Guid("1b27fa6c-fa0f-e911-a94e-000d3af060a1"))//Part PO Created
            {
                Entity workOrder = new Entity("msdyn_workorder", workOrderRef.Id);
                workOrder["msdyn_substatus"] = workInitiated;
                service.Update(workOrder);
            }
            else
            {
                Console.WriteLine($"Job Status: {_jobSubStatus.Name}");
            }
        }

        private EntityReference? GetFranchise(string franchiseCode)
        {
            franchiseCode = franchiseCode[..1] == "F" ? franchiseCode[1..] : franchiseCode;
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

        private Guid GetInventoryPurchaseRecept(string sapInvoiceNumber, EntityReference franchise, EntityReference purchaseOrder, EntityReference freshWarehouse, Guid ownerId, EntityReference workOrderRef)
        {
            QueryExpression query = new QueryExpression("hil_inventorypurchaseorderreceipt");
            query.Criteria.AddCondition("hil_invoicenumber", ConditionOperator.Equal, sapInvoiceNumber);
            query.Criteria.AddCondition("hil_receiptstatus", ConditionOperator.Equal, 1); // draft

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
            if (workOrderRef != null) newPurchaseReceipt["hil_jobid"] = workOrderRef;
            return service.Create(newPurchaseReceipt);
        }

        private Entity GetOrderLine(EntityReference product, EntityReference purhaseOrder)
        {
            QueryExpression query = new QueryExpression("hil_inventorypurchaseorderline");
            query.Criteria.AddCondition("hil_partcode", ConditionOperator.Equal, product.Id);
            query.Criteria.AddCondition("hil_ponumber", ConditionOperator.Equal, purhaseOrder.Id);

            EntityCollection orderLineColl = service.RetrieveMultiple(query);
            return orderLineColl.Entities[0];
        }

        private Guid GetPurchaseOrderReceiptLine(Guid purchaseOrderReceipt, EntityReference product, EntityReference purchaseOrder, EntityReference freshWarehouse, Guid ownerId, EntityReference workOrderRef, int billedQuantity)
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
            if (workOrderRef != null) newPurchaseLine["hil_jobid"] = workOrderRef;
            newPurchaseLine["hil_billedquantity"] = billedQuantity;
            return service.Create(newPurchaseLine);

        }

        internal void GetSapInvoicetoSync(string _startDate, string _endDate)
        {
            IntegrationConfiguration integrationConfiguration = HelperClass.GetIntegrationConfiguration(service, "CRM_SAPtoCRM_Invoice");
            //DateTime FromDate = DateTime.Now.AddDays(-1);  //GetLastRunDate(service);
            //String fromDate = FromDate.Year.ToString() + FromDate.Month.ToString().PadLeft(2, '0') + FromDate.Day.ToString().PadLeft(2, '0');

            //DateTime ToDate = DateTime.Now;
            //String toDate = ToDate.Year.ToString() + ToDate.Month.ToString().PadLeft(2, '0') + ToDate.Day.ToString().PadLeft(2, '0');

            //integrationConfiguration.url = "https://middlewareqa.havells.com:50001/RESTAdapter/dynamics/invoice?IM_FLAG=R&IM_PROJECT=SER";
            integrationConfiguration.url = integrationConfiguration.url.Replace("middleware", "p90ci");

            integrationConfiguration.url = integrationConfiguration.url + $"&IM_FROM_DT={_startDate}&IM_TO_DT={_endDate}";

            //integrationConfiguration.url = integrationConfiguration.url + "&IM_FROM_DT=" + fromDate + "&IM_TO_DT=" + toDate;

            //integrationConfiguration.url = integrationConfiguration.url + "&IM_FROM_DT=" + "20240110" + "&IM_TO_DT=" + "20240130";
            //integrationConfiguration.url = integrationConfiguration.url.Replace("middleware", "p90ci");
            //integrationConfiguration.url = integrationConfiguration.url.Replace("middleware", "p90ci");

            //integrationConfiguration.password = "QAD365@1234";

            InvoiceResonse resp = JsonConvert.DeserializeObject<InvoiceResonse>(HelperClass.CallAPI(integrationConfiguration, null, "GET"));

            if (resp.LT_INVOICE.Count > 0)
            {
                QueryExpression query;
                int recNo = 1;
                foreach (LTINVOICE lTINVOICE in resp.LT_INVOICE)
                {
                    try
                    {
                        Console.WriteLine($"Record# {recNo++}/{resp.LT_INVOICE.Count} Invoice# {lTINVOICE.VBELN} SO# {lTINVOICE.AUBEL}");
                        if (lTINVOICE != null)
                        {
                            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_inventorypurchaseorder'>
                                <attribute name='hil_inventorypurchaseorderid' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                    <condition attribute='hil_salesordernumber' operator='eq' value='{lTINVOICE.AUBEL.Trim()}' />
                                </filter>
                                </entity>
                                </fetch>";
                            EntityCollection entColPO = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entColPO.Entities.Count > 0)
                            {
                                string franchiseCode = lTINVOICE.KUNRG.Substring(1);

                                EntityReference product = GetProduct(lTINVOICE.MATNR); // lTINVOICE.MATNR
                                EntityReference franchise = GetFranchise(lTINVOICE.KUNRG); // lTINVOICE.KUNRG

                                if (product != null && franchise != null)
                                {
                                    query = new QueryExpression("hil_inventorypurchaseorder");
                                    query.ColumnSet = new ColumnSet("ownerid", "hil_jobid");
                                    query.Criteria.AddCondition("hil_salesordernumber", ConditionOperator.Equal, lTINVOICE.AUBEL);
                                    EntityCollection purchaseOrderCollection = service.RetrieveMultiple(query);

                                    query = new QueryExpression("hil_inventorywarehouse");
                                    query.Criteria.AddCondition("hil_franchise", ConditionOperator.Equal, franchise.Id);
                                    query.Criteria.AddCondition("hil_type", ConditionOperator.Equal, 1); //fresh 
                                    EntityCollection wareHouseCollection = service.RetrieveMultiple(query);

                                    if (purchaseOrderCollection.Entities.Count > 0 && wareHouseCollection.Entities.Count > 0)
                                    {
                                        query = new QueryExpression("hil_inventorysparebills");
                                        query.Criteria.AddCondition("hil_salesordernumber", ConditionOperator.Equal, lTINVOICE.AUBEL);
                                        query.Criteria.AddCondition("hil_billnumber", ConditionOperator.Equal, lTINVOICE.VBELN);
                                        query.Criteria.AddCondition("hil_billdate", ConditionOperator.Equal, lTINVOICE.FKDAT);
                                        query.Criteria.AddCondition("hil_franchise", ConditionOperator.Equal, franchise.Id);
                                        query.Criteria.AddCondition("hil_partcode", ConditionOperator.Equal, product.Id);
                                        EntityCollection doesBillExist = service.RetrieveMultiple(query);
                                        if (doesBillExist.Entities.Count == 0)
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
                                        }

                                        EntityReference workOrderRef = purchaseOrderCollection.Entities[0].Contains("hil_jobid") ? purchaseOrderCollection.Entities[0].GetAttributeValue<EntityReference>("hil_jobid") : null;

                                        EntityReference poOwner = purchaseOrderCollection.Entities[0].GetAttributeValue<EntityReference>("ownerid");

                                        Guid purchaseOrderReceipt = GetInventoryPurchaseRecept(lTINVOICE.VBELN, franchise, purchaseOrderCollection.Entities[0].ToEntityReference(), wareHouseCollection.Entities[0].ToEntityReference(), poOwner.Id, workOrderRef);

                                        Entity _entGRN = service.Retrieve("hil_inventorypurchaseorderreceipt", purchaseOrderReceipt, new ColumnSet("hil_receiptstatus"));
                                        if (_entGRN != null)
                                        {
                                            OptionSetValue _grnStatus = _entGRN.GetAttributeValue<OptionSetValue>("hil_receiptstatus");
                                            if (_grnStatus.Value == 1)//GRN : DRAFT
                                            {
                                                int billedQuantity = Convert.ToInt32(Convert.ToDecimal(lTINVOICE.FKIMG));//Quantity

                                                Guid purchaseOrderReceiptLine = GetPurchaseOrderReceiptLine(purchaseOrderReceipt, product, purchaseOrderCollection.Entities[0].ToEntityReference(), wareHouseCollection.Entities[0].ToEntityReference(), poOwner.Id, workOrderRef, billedQuantity);

                                                Entity updateLine = new Entity("hil_inventorypurchaseorderreceiptline", purchaseOrderReceiptLine);

                                                service.Update(updateLine);
                                                //if (workOrderRef != null)
                                                //UpdateWorkOrder(workOrderRef);
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine($"SO# {lTINVOICE.AUBEL.Trim()} doesn't exist in CRM.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("error " + ex.Message);
                    }
                }
            }
            //ReCalSuppliedQuantityAndRefreshJobSubstatus();
        }

        internal void ReCalSuppliedQuantityAndRefreshJobSubstatus()
        {
            string fetchPurchaseOrders = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='hil_inventorypurchaseorder'>
            <attribute name='hil_inventorypurchaseorderid'/>
            <attribute name='hil_name'/>
            <attribute name='createdon'/>
            <attribute name='hil_jobid' />
            <attribute name='hil_salesordernumber'/>
            <order attribute='hil_name' descending='false'/>
            <filter type='and'>
            <condition attribute='createdon' operator='on-or-after' value='2024-12-12'/>
            <condition attribute='hil_salesordernumber' operator='not-null'/>
            </filter>
            <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='hil_jobid' link-type='inner' alias='aw'>
            <filter type='and'>
            <condition attribute='msdyn_substatus' operator='eq' uiname='Part PO Created' uitype='msdyn_workordersubstatus' value='{1B27FA6C-FA0F-E911-A94E-000D3AF060A1}'/>
            </filter>
            </link-entity>
            </entity>
            </fetch>";

            EntityCollection pendingPurchaseOrders = service.RetrieveMultiple(new FetchExpression(fetchPurchaseOrders));

            foreach (Entity purchaseOrder in pendingPurchaseOrders.Entities)
            {
                string salesOrderNumber = purchaseOrder.GetAttributeValue<string>("hil_salesordernumber");

                string fetchTotalBilledQuantity = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' aggregate='true'>
                <entity name='hil_inventorypurchaseorderreceiptline'>
                <attribute name='hil_purchaseorderline' alias='poline' groupby='true'/>
                <attribute name='hil_billedquantity' alias='billqty' aggregate='sum'/>
                <filter type='and'>
                <condition attribute='hil_ordernumber' operator='eq' value='{purchaseOrder.Id}' />
                </filter>
                </entity>
                </fetch>";
                EntityCollection tempColl = service.RetrieveMultiple(new FetchExpression(fetchTotalBilledQuantity));
                if (tempColl.Entities.Count > 0)
                {
                    foreach (var entity in tempColl.Entities)
                    {
                        int _totalBillQty = (int)entity.GetAttributeValue<AliasedValue>("billqty").Value;

                        EntityReference _poline = (EntityReference)entity.GetAttributeValue<AliasedValue>("poline").Value;

                        Entity _entUpdate = new Entity(_poline.LogicalName, _poline.Id);

                        _entUpdate["hil_suppliedquantity"] = _totalBillQty;

                        service.Update(_entUpdate);

                        Entity purchaseOrderLine = service.Retrieve(_poline.LogicalName, _poline.Id, new ColumnSet("hil_pendingquantity", "hil_partstatus", "hil_suppliedquantity"));
                        int pendingQuantity = purchaseOrderLine.GetAttributeValue<int>("hil_pendingquantity");
                        int suppliedQuantity = purchaseOrderLine.GetAttributeValue<int>("hil_suppliedquantity");

                        if (suppliedQuantity > 0 && pendingQuantity <= 0)
                        {
                            purchaseOrderLine["hil_partstatus"] = new OptionSetValue(3); //Dispatched
                        }
                        else if (suppliedQuantity > 0 && pendingQuantity != 0)
                        {
                            purchaseOrderLine["hil_partstatus"] = new OptionSetValue(2); //Partially Dispatched
                        }
                        service.Update(purchaseOrderLine);

                        if (purchaseOrder.Contains("hil_jobid"))
                        {
                            EntityReference jobId = purchaseOrder.GetAttributeValue<EntityReference>("hil_jobid");
                            Entity woObj = service.Retrieve(jobId.LogicalName, jobId.Id, new ColumnSet("msdyn_substatus"));
                            EntityReference _jobSubStatus = woObj.GetAttributeValue<EntityReference>("msdyn_substatus");

                            if (_jobSubStatus.Id == new Guid("1b27fa6c-fa0f-e911-a94e-000d3af060a1"))//Part PO Created
                            {
                                string fetchExpression = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                    <entity name=""hil_inventorypurchaseorderline"">
                                    <attribute name=""hil_inventorypurchaseorderlineid""/>
                                    <attribute name=""hil_name""/>
                                    <attribute name=""createdon""/>
                                    <order attribute=""hil_name"" descending=""false""/>
                                    <filter type=""and"">
                                    <condition attribute=""hil_workorder"" operator=""eq"" uiname=""31122326855541"" uitype=""msdyn_workorder"" value=""{woObj.Id}""/>
                                    <condition attribute=""hil_partstatus"" operator=""not-in"">
                                    <value>3</value>
                                    <value>5</value>
                                    </condition>
                                    </filter>
                                    </entity>
                                    </fetch>";

                                EntityCollection retriveJobs = service.RetrieveMultiple(new FetchExpression(fetchExpression));

                                if (retriveJobs.Entities.Count == 0)
                                {
                                    Entity workOrder = new Entity("msdyn_workorder", woObj.Id);
                                    workOrder["msdyn_substatus"] = workInitiated;
                                    service.Update(workOrder);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}