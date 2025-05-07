using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Web.Services.Description;

namespace Havells.CRM.DynamicsPOSyncToSAP
{
    public class SyncInvoiceSAPtoD365
    {
        public static void GetSapInvoicetoSync(IOrganizationService service)
        {
            IntegrationConfiguration integrationConfiguration = HelperClass.GetIntegrationConfiguration(service, "CRM_SAPtoCRM_Invoice");
            DateTime FromDate = DateTime.Now.AddDays(-1);  //GetLastRunDate(service);
            String ToDate = FromDate.Year.ToString() + FromDate.Month.ToString().PadLeft(2, '0') + FromDate.Day.ToString().PadLeft(2, '0');

            integrationConfiguration.url = integrationConfiguration.url + "&IM_FROM_DT=" + ToDate + "&IM_TO_DT=" + ToDate;

            InvoiceResonse resp = JsonConvert.DeserializeObject<InvoiceResonse>(HelperClass.CallAPI(integrationConfiguration, null, "GET"));

            if (resp.LT_INVOICE.Count > 0)
            {
                foreach (LTINVOICE lTINVOICE in resp.LT_INVOICE)
                {
                    try
                    {
                        if (lTINVOICE != null)
                        {
                            Entity invoiceDeatils = new Entity("hil_invoice");
                            invoiceDeatils["hil_name"] = lTINVOICE.VBELN;
                            invoiceDeatils["hil_companyname"] = lTINVOICE.BUTXT;

                            invoiceDeatils["hil_cess"] = Math.Round(Convert.ToDecimal(lTINVOICE.CESS_PERC), 2);

                            invoiceDeatils["hil_cgst"] = Math.Round(Convert.ToDecimal(lTINVOICE.CGST_PERC), 2);
                            invoiceDeatils["hil_cgstvalue"] = Math.Round(Convert.ToDecimal(lTINVOICE.CGST_AMT), 2);

                            invoiceDeatils["hil_dlp"] = Math.Round(Convert.ToDecimal(lTINVOICE.DLP), 2);
                            invoiceDeatils["hil_igst"] = Math.Round(Convert.ToDecimal(lTINVOICE.IGST_PERC), 2);

                            invoiceDeatils["hil_igstvalue"] = Math.Round(Convert.ToDecimal(lTINVOICE.IGST_AMT), 2);

                            invoiceDeatils["hil_invoicedate"] = lTINVOICE.FKDAT;

                            invoiceDeatils["hil_invoicenetamount"] = Math.Round(Convert.ToDecimal(lTINVOICE.NETWR), 2);
                            invoiceDeatils["hil_lldiscount"] = Math.Round(Convert.ToDecimal(lTINVOICE.LLDISCOUNT), 2);
                            invoiceDeatils["hil_llamount"] = Math.Round(Convert.ToDecimal(lTINVOICE.LLAMOUNT), 2);
                            invoiceDeatils["hil_lltaxamt"] = Math.Round(Convert.ToDecimal(lTINVOICE.LLTAXAMT), 2);
                            invoiceDeatils["hil_lltotalamt"] = Math.Round(Convert.ToDecimal(lTINVOICE.LLTOTAMT), 2);
                            invoiceDeatils["hil_mrp"] = Math.Round(Convert.ToDecimal(lTINVOICE.MRP), 2);

                            invoiceDeatils["hil_productcode"] = lTINVOICE.MATNR;
                            invoiceDeatils["hil_purchaserate"] = Math.Round(Convert.ToDecimal(lTINVOICE.PUR_RATE), 2);
                            if (!string.IsNullOrEmpty(lTINVOICE.FKIMG))
                            {
                                invoiceDeatils["hil_quantity"] = Convert.ToInt32(Convert.ToDecimal(lTINVOICE.FKIMG));
                            }
                            invoiceDeatils["hil_salesordernumber"] = lTINVOICE.AUBEL;
                            invoiceDeatils["hil_sgst"] = Math.Round(Convert.ToDecimal(lTINVOICE.SGST_PERC), 2);
                            invoiceDeatils["hil_sgstvalue"] = Math.Round(Convert.ToDecimal(lTINVOICE.SGST_AMT), 2);
                            invoiceDeatils["hil_stockistcode"] = lTINVOICE.KUNRG;
                            invoiceDeatils["hil_suppliercode"] = lTINVOICE.WERKS;
                            invoiceDeatils["hil_taxablevalue"] = Math.Round(Convert.ToDecimal(lTINVOICE.TAXABLE_AMT), 2);
                            invoiceDeatils["hil_totaldiscount"] = Math.Round(Convert.ToDecimal(lTINVOICE.TOT_DISC), 2);
                            invoiceDeatils["hil_totaltax"] = Math.Round(Convert.ToDecimal(lTINVOICE.MWSBK), 2);
                            invoiceDeatils["hil_transporter"] = lTINVOICE.BEZEI;
                            invoiceDeatils["hil_uom"] = lTINVOICE.MEINS;
                            invoiceDeatils["hil_utgst"] = Math.Round(Convert.ToDecimal(lTINVOICE.UTGST_PERC), 2);
                            invoiceDeatils["hil_utgstvalue"] = Math.Round(Convert.ToDecimal(lTINVOICE.UTGST_AMT), 2);
                            service.Create(invoiceDeatils);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("error " + ex.Message);
                    }
                }
            }
        }
        static string GetLastRunDate(IOrganizationService service)
        {
            string _lastAMCPurchaseDate = "2024-04-28";
            try
            {
                QueryExpression Query = new QueryExpression("hil_invoice");
                Query.ColumnSet = new ColumnSet("modifiedon", "createdon");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("modifiedon", ConditionOperator.NotNull);
                Query.TopCount = 1;
                Query.AddOrder("modifiedon", OrderType.Descending);

                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    DateTime _date = Found.Entities[0].GetAttributeValue<DateTime>("modifiedon").AddMinutes(330);
                    _lastAMCPurchaseDate = _date.Year.ToString() + "-" + _date.Month.ToString().PadLeft(2, '0') + "-" + _date.Day.ToString().PadLeft(2, '0');
                }
                return _lastAMCPurchaseDate;
            }
            catch { }
            return _lastAMCPurchaseDate;
        }

        //static EntityReference RetriveRecipt(IOrganizationService service, string ReciptNumber)
        //{
        //    QueryExpression qsCType = new QueryExpression("msdyn_purchaseorderreceipt");
        //    qsCType.ColumnSet = new ColumnSet(false);
        //    qsCType.NoLock = true;
        //    qsCType.Criteria = new FilterExpression(LogicalOperator.And);
        //    qsCType.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, ReciptNumber);
        //    EntityCollection entity = service.RetrieveMultiple(qsCType);
        //    if (entity.Entities.Count == 1)
        //        return entity[0].ToEntityReference();
        //    else
        //        return null;
        //}
        //static Entity RetrivePO(IOrganizationService service, string SalesOrderNumber)
        //{
        //    QueryExpression qsCType = new QueryExpression("msdyn_purchaseorder");
        //    qsCType.ColumnSet = new ColumnSet("msdyn_receivetowarehouse", "msdyn_workorder", "ownerid");
        //    qsCType.NoLock = true;
        //    qsCType.Criteria = new FilterExpression(LogicalOperator.And);
        //    qsCType.Criteria.AddCondition("hil_sapsalesorderno", ConditionOperator.Equal, SalesOrderNumber);
        //    EntityCollection entity = service.RetrieveMultiple(qsCType);
        //    if (entity.Entities.Count == 1)
        //        return entity[0];
        //    else
        //        return null;
        //}
        //static EntityReference RetrivePOPorduct(IOrganizationService service, string productName, Guid POID, string quantity)
        //{
        //    string fetch = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""true"">
        //              <entity name=""msdyn_purchaseorderproduct"">
        //                <attribute name=""createdon"" />
        //                <attribute name=""msdyn_quantity"" />
        //                <attribute name=""msdyn_purchaseorder"" />
        //                <attribute name=""msdyn_product"" />
        //                <attribute name=""msdyn_purchaseorderproductid"" />
        //                <order attribute=""msdyn_purchaseorder"" descending=""true"" />
        //                <filter type=""and"">
        //                  <condition attribute=""msdyn_quantity"" operator=""eq"" value=""{quantity}"" />
        //                  <condition attribute=""msdyn_purchaseorder"" operator=""eq"" value=""{POID}"" />
        //                  <condition attribute=""statecode"" operator=""eq"" value=""0"" />
        //                </filter>
        //                <link-entity name=""msdyn_purchaseorderreceiptproduct"" from=""msdyn_purchaseorderproduct"" to=""msdyn_purchaseorderproductid"" link-type=""outer"" alias=""ai"" />
        //                <link-entity name=""product"" from=""productid"" to=""msdyn_product"" link-type=""inner"" alias=""product"">
        //                  <filter type=""and"">
        //                    <condition attribute=""name"" operator=""eq"" value=""{productName}"" />
        //                  </filter>
        //                </link-entity>
        //                <filter type=""and"">
        //                  <condition entityname=""ai"" attribute=""msdyn_purchaseorderproduct"" operator=""null"" />
        //                </filter>
        //              </entity>
        //            </fetch>";
        //    EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetch));
        //    if (entityCollection.Entities.Count != 0)
        //        return entityCollection.Entities[0].ToEntityReference();
        //    else
        //        return null;

        //}
        //static void CreateReciptProduct(IOrganizationService service, LTINVOICE lTINVOICE, Entity PO, EntityReference Recipt, EntityReference PO_ProductRef)
        //{
        //    Entity reciptProductEntity = new Entity("msdyn_purchaseorderreceiptproduct");
        //    reciptProductEntity["msdyn_name"] = lTINVOICE.VBELN + "_" + int.Parse(lTINVOICE.POSNR).ToString();
        //    reciptProductEntity["msdyn_purchaseorderreceipt"] = Recipt;
        //    reciptProductEntity["ownerid"] = PO["ownerid"];
        //    reciptProductEntity["msdyn_quantity"] = 0.00;
        //    reciptProductEntity["msdyn_purchaseorder"] = PO.ToEntityReference();
        //    reciptProductEntity["hil_billedquantity"] = Convert.ToInt32(Convert.ToDecimal(lTINVOICE.FKIMG));
        //    reciptProductEntity["msdyn_associatetoworkorder"] = PO["msdyn_workorder"];
        //    reciptProductEntity["msdyn_associatetowarehouse"] = PO["msdyn_receivetowarehouse"];
        //    reciptProductEntity["msdyn_purchaseorderproduct"] = PO_ProductRef;
        //    reciptProductEntity["msdyn_unitcost"] = new Money(Convert.ToDecimal(lTINVOICE.PUR_RATE));
        //    reciptProductEntity["msdyn_totalcost"] = new Money(Convert.ToDecimal(lTINVOICE.LLTOTAMT));
        //    service.Create(reciptProductEntity);
        //}

    }
}
