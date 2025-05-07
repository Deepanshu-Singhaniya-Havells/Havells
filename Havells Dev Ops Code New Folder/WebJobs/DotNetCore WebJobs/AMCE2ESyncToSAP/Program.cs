using Microsoft.PowerPlatform.Dataverse.Client;
using System.Configuration;

namespace AMCE2ESyncToSAP
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"];
            var CrmURLDev = ConfigurationManager.AppSettings["CRMUrl"];
            string finalConnStr = string.Format(connStr ?? "", CrmURLDev);
            ServiceClient service = new ServiceClient(finalConnStr);

            ClsAMCE2E obj = new ClsAMCE2E(service);
           // obj.getPaymentStatusFromPayU();
           // obj.PushAMCDataJob();
            obj.PushAMCData_OmniChannel();
        }
    }
}
