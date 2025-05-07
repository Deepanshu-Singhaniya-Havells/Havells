using Microsoft.PowerPlatform.Dataverse.Client;
using SOPaymentReceipt;
using System.Configuration;

namespace SOSyncToSAP
{
    public class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"];
            var CrmURLDev = ConfigurationManager.AppSettings["CRMUrl"];
            string finalConnStr = string.Format(connStr ?? "", CrmURLDev);
            ServiceClient service = new ServiceClient(finalConnStr);
            Guid SalesOrderId = string.IsNullOrWhiteSpace(args[0]) ? Guid.Empty : new Guid(args[0].ToString());
            ClsSalesOrder obj = new ClsSalesOrder(service);
            obj.RePushSalesOrderToSAP(SalesOrderId);
        }
    }
}
