using KKGCloseJobsForMoengage;
using Microsoft.PowerPlatform.Dataverse.Client;
using Newtonsoft.Json;
using System.Configuration;

namespace InventoryWebJobs
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
          
            var connStr = ConfigurationManager.AppSettings["connStr"];
            var CrmURLProd = ConfigurationManager.AppSettings["CRMUrlProd"];
            string finalConnStr = string.Format(connStr ?? "", CrmURLProd);
            ServiceClient service = new ServiceClient(finalConnStr);
            ClsKKGCloseJobsData obj = new ClsKKGCloseJobsData(service);
            obj.ClsKKGCloseJobs();
        }   
 }
}
