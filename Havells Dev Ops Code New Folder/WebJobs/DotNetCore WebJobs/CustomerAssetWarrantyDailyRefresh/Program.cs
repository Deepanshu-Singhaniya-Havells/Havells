using Microsoft.PowerPlatform.Dataverse.Client;
using System.Configuration;

namespace CustomerAssetWarrantyDailyRefresh
{
    public class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"];
            var CrmURLDev = ConfigurationManager.AppSettings["CRMUrl"];
            string finalConnStr = string.Format(connStr ?? "", CrmURLDev);
            ServiceClient service = new ServiceClient(finalConnStr);

            RefreshUnitWarranty refreshUnitWarranty = new RefreshUnitWarranty(service);
            //refreshUnitWarranty.RefreshAssetUnitWarranty();
            refreshUnitWarranty.RefreshAssetAMCUnitWarranty();
        }
    }
}
