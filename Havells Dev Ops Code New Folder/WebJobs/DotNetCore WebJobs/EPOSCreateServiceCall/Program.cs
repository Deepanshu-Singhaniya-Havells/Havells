using Microsoft.PowerPlatform.Dataverse.Client;
using System.Configuration;

namespace EPOSCreateServiceCall
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"];
            var CrmURLDev = ConfigurationManager.AppSettings["CRMUrl"];
            string finalConnStr = string.Format(connStr ?? "", CrmURLDev);
            ServiceClient service = new ServiceClient(finalConnStr);

            ClsCreateJobs obj = new ClsCreateJobs(service);
            obj.getAllActiveSFAServiceRequest();
        }
    }
}
