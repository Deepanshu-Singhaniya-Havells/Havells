using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Configuration;

namespace Havells.CRM.GetChannelPartner
{
    class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = ConnectToCRM(finalString);
            if (((CrmServiceClient)service).IsReady)
                //GetRetailersMaster.getRetailersMaster(service, (args.Length > 0 ? args[0] : ""));
                GetPartner.GetChannelPartnerData(service, "");
        }
        public static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
    }
}
