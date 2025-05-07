using Microsoft.PowerPlatform.Dataverse.Client;
using System.Configuration;

namespace PaymentStatus
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"];
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"];
            string finalConnStr = string.Format(connStr ?? "", CrmURL);
            ServiceClient service = new ServiceClient(finalConnStr);

            ClsSOUpdatePaymentStatus obj = new ClsSOUpdatePaymentStatus(service);
            obj.UpdatePaymentStatus();
        }
    }
}
