using Microsoft.PowerPlatform.Dataverse.Client;
using System.Configuration;

namespace ChannelPartnerSync
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"];
            var CrmURLProd = ConfigurationManager.AppSettings["CRMUrlProd"];
            string finalConnStr = string.Format(connStr ?? "", CrmURLProd);
            ServiceClient service = new ServiceClient(finalConnStr);

            ClsPartner obj = new ClsPartner(service);
            obj.GetChannelPartnerData(args.Length > 0 ? args[0] : "");
        }
    }
}
