using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;

namespace Havells.CRM.DynamicsPOSyncToSAP
{
    public class Sync_Invoice_SAP_to_D365
    {
        public static void GetAllInvoice(IOrganizationService service)
        {
            IntegrationConfiguration integrationConfiguration = HelperClass.GetIntegrationConfiguration(service, "CRM_SAPtoCRM_Invoice");
            InvoiceResonse resp = JsonConvert.DeserializeObject<InvoiceResonse>(HelperClass.CallAPI(integrationConfiguration, null, "GET"));
            string invoiceID_VBELN = string.Empty;
            string SalesOrderNumber_AUBEL = string.Empty;
            EntityReference POReciptRef = null;
            EntityReference po_ProductRef = null;
            Entity PO_Entity = null;
            foreach (LTINVOICE lTINVOICE in resp.LT_INVOICE)
            {
                try
                {
                    invoiceID_VBELN = lTINVOICE.VBELN;
                    SalesOrderNumber_AUBEL = lTINVOICE.AUBEL;
                    POReciptRef = RetriveRecipt(service, invoiceID_VBELN);
                    PO_Entity = RetrivePO(service, SalesOrderNumber_AUBEL);
                    if (PO_Entity != null)
                        po_ProductRef = RetrivePOPorduct(service, lTINVOICE.MATNR, PO_Entity.Id, lTINVOICE.FKIMG);
                    if (PO_Entity != null && po_ProductRef != null && POReciptRef != null)
                    {
                        CreateReciptProduct(service, lTINVOICE, PO_Entity, POReciptRef, po_ProductRef);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("error " + ex.Message);
                }
            }
        }
        static EntityReference RetriveRecipt(IOrganizationService service, string ReciptNumber)
        {
            QueryExpression qsCType = new QueryExpression("msdyn_purchaseorderreceipt");
            qsCType.ColumnSet = new ColumnSet(false);
            qsCType.NoLock = true;
            qsCType.Criteria = new FilterExpression(LogicalOperator.And);
            qsCType.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, ReciptNumber);
            EntityCollection entity = service.RetrieveMultiple(qsCType);
            if (entity.Entities.Count == 1)
                return entity[0].ToEntityReference();
            else
                return null;
        }
        static Entity RetrivePO(IOrganizationService service, string SalesOrderNumber)
        {
            QueryExpression qsCType = new QueryExpression("msdyn_purchaseorder");
            qsCType.ColumnSet = new ColumnSet("msdyn_receivetowarehouse", "msdyn_workorder", "ownerid");
            qsCType.NoLock = true;
            qsCType.Criteria = new FilterExpression(LogicalOperator.And);
            qsCType.Criteria.AddCondition("hil_sapsalesorderno", ConditionOperator.Equal, SalesOrderNumber);
            EntityCollection entity = service.RetrieveMultiple(qsCType);
            if (entity.Entities.Count == 1)
                return entity[0];
            else
                return null;
        }
        static EntityReference RetrivePOPorduct(IOrganizationService service, string productName, Guid POID,string quantity)
        {
            string fetch = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""true"">
                      <entity name=""msdyn_purchaseorderproduct"">
                        <attribute name=""createdon"" />
                        <attribute name=""msdyn_quantity"" />
                        <attribute name=""msdyn_purchaseorder"" />
                        <attribute name=""msdyn_product"" />
                        <attribute name=""msdyn_purchaseorderproductid"" />
                        <order attribute=""msdyn_purchaseorder"" descending=""true"" />
                        <filter type=""and"">
                          <condition attribute=""msdyn_quantity"" operator=""eq"" value=""{quantity}"" />
                          <condition attribute=""msdyn_purchaseorder"" operator=""eq"" value=""{POID}"" />
                          <condition attribute=""statecode"" operator=""eq"" value=""0"" />
                        </filter>
                        <link-entity name=""msdyn_purchaseorderreceiptproduct"" from=""msdyn_purchaseorderproduct"" to=""msdyn_purchaseorderproductid"" link-type=""outer"" alias=""ai"" />
                        <link-entity name=""product"" from=""productid"" to=""msdyn_product"" link-type=""inner"" alias=""product"">
                          <filter type=""and"">
                            <condition attribute=""name"" operator=""eq"" value=""{productName}"" />
                          </filter>
                        </link-entity>
                        <filter type=""and"">
                          <condition entityname=""ai"" attribute=""msdyn_purchaseorderproduct"" operator=""null"" />
                        </filter>
                      </entity>
                    </fetch>";
            EntityCollection entityCollection = service.RetrieveMultiple(new FetchExpression(fetch));
            if (entityCollection.Entities.Count != 0)
                return entityCollection.Entities[0].ToEntityReference();
            else
                return null;

        }
        static void CreateReciptProduct(IOrganizationService service, LTINVOICE lTINVOICE, Entity PO, EntityReference Recipt, EntityReference PO_ProductRef)
        {
            Entity reciptProductEntity = new Entity("msdyn_purchaseorderreceiptproduct");
            reciptProductEntity["msdyn_name"] = lTINVOICE.VBELN + "_" + int.Parse(lTINVOICE.POSNR).ToString();
            reciptProductEntity["msdyn_purchaseorderreceipt"] = Recipt;
            reciptProductEntity["ownerid"] = PO["ownerid"];
            reciptProductEntity["msdyn_quantity"] = 0.00;
            reciptProductEntity["msdyn_purchaseorder"] = PO.ToEntityReference();
            reciptProductEntity["hil_billedquantity"] = Convert.ToInt32(Convert.ToDecimal(lTINVOICE.FKIMG));
            reciptProductEntity["msdyn_associatetoworkorder"] = PO["msdyn_workorder"];
            reciptProductEntity["msdyn_associatetowarehouse"] = PO["msdyn_receivetowarehouse"];
            reciptProductEntity["msdyn_purchaseorderproduct"] = PO_ProductRef;
            reciptProductEntity["msdyn_unitcost"] = new Money(Convert.ToDecimal(lTINVOICE.PUR_RATE));
            reciptProductEntity["msdyn_totalcost"] = new Money(Convert.ToDecimal(lTINVOICE.LLTOTAMT));
            service.Create(reciptProductEntity);
        }
    }
}
