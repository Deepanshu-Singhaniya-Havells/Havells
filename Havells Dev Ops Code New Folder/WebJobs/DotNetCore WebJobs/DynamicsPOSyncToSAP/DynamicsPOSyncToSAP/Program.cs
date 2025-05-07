using System.Configuration;
using Microsoft.PowerPlatform.Dataverse.Client;

namespace DynamicsPOSyncToSAP
{
	static class Program
    {       
        static void Main(string[] args)
        {
			var connStr = ConfigurationManager.AppSettings["connStr"];
			var CrmURLDev = ConfigurationManager.AppSettings["CRMUrl"];
			string finalConnStr = string.Format(connStr ?? "", CrmURLDev);
			ServiceClient service = new ServiceClient(finalConnStr);

			SyncSAPtoD365BillingDetail syncSAPtoD365BillingDetail = new SyncSAPtoD365BillingDetail(service);
			syncSAPtoD365BillingDetail.UpdateAMCStagingTable();
        }
    }
}