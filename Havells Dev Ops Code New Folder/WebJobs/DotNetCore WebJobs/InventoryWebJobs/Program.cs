using InventoryWebJobs.ProductReplacement;
using Microsoft.PowerPlatform.Dataverse.Client;
using System.Configuration;

namespace InventoryWebJobs
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"];
            var CrmURLProd = ConfigurationManager.AppSettings["CRMUrlProd"];
            var StartDate = ConfigurationManager.AppSettings["StartDate"];
            var EndDate = ConfigurationManager.AppSettings["EndDate"];
            string finalConnStr = string.Format(connStr ?? "", CrmURLProd);
            ServiceClient service = new ServiceClient(finalConnStr);

            POtoSAP pOtoSAPObj = new POtoSAP(service);
            pOtoSAPObj.SyncPoToSAP();

            //SyncInvoice syncInvoice = new SyncInvoice(service);
            //syncInvoice.GetSapInvoicetoSync(StartDate, EndDate);
            //syncInvoice.ReCalSuppliedQuantityAndRefreshJobSubstatus();

            //PostRMA ObjPostRMA = new PostRMA(service);
            //await ObjPostRMA.PostRMAs();

            //SyncProductReplacement syncProductreplacements = new SyncProductReplacement(service);
            //syncProductreplacements.GetProductRequisition();
        }
    }
}
