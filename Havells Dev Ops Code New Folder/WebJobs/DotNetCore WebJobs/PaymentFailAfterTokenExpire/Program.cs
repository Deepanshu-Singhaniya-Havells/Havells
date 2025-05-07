using Microsoft.PowerPlatform.Dataverse.Client;
using System.Configuration;

namespace PaymentFailAfterTokenExpire
{
    public class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"];
            var CrmURLProd = ConfigurationManager.AppSettings["CRMUrlProd"];
            string finalConnStr = string.Format(connStr ?? "", CrmURLProd);
            ServiceClient service = new ServiceClient(finalConnStr);

            MakePaymentFailed Obj = new MakePaymentFailed(service);
            Obj.UpdateStatusAfterTokenExpire();

        }
    }
}
