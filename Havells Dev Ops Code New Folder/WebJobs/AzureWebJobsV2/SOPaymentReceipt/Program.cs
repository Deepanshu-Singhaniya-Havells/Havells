using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;

namespace SOPaymentReceipt
{
	public class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        #region Global Varialble declaration
        static IOrganizationService _service;
        #endregion
        static void Main(string[] args)
        {
            try
            {
                _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
                if (((CrmServiceClient)_service).IsReady)
                {
                    //ClsPaymentReceipt obj = new ClsPaymentReceipt(_service);
                    //obj.PaymentReceipt();

                    ClsSalesOrder obj1 = new ClsSalesOrder(_service);
                    obj1.PushSalesOrderToSAP();
                }
            }
            catch (Exception ex)
            {
                Console.Write("ERROR!!! " + ex.Message);
            }
        }
        #region App Setting Load/CRM Connection
        static IOrganizationService ConnectToCRM(string connectionString)
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
        #endregion
    }
}
