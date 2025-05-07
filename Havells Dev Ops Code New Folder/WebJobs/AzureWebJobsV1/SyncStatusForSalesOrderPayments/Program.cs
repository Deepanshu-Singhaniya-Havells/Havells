using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;


namespace SyncStatusForSalesOrderPayments
{
    internal class Program
    {


        private static IOrganizationService service; 


        static void Main(string[] args)
        {
            service = ConnectToCRM(); 
            SyncPayments syncPaymentsObj = new SyncPayments(service);
            syncPaymentsObj.FetchPayments();



        }





        private static IOrganizationService ConnectToCRM()
        {

            string devURL = "https://havellscrmdev1.crm8.dynamics.com";
            string connectionString = $"AuthType=ClientSecret;url={devURL};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

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
