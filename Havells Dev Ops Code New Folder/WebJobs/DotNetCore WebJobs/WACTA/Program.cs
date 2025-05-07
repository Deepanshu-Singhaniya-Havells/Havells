using Microsoft.PowerPlatform.Dataverse.Client;
using System.Configuration;

namespace WACTA
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"];
            var CrmURLDev = ConfigurationManager.AppSettings["CRMUrl"];
            var triggerName = ConfigurationManager.AppSettings["triggerName"];
            string finalConnStr = string.Format(connStr ?? "", CrmURLDev);
            ServiceClient service = new ServiceClient(finalConnStr);

            ClsWACTA obj = new ClsWACTA(service);
            try
            {
                obj.SentDataForNewWACTA(triggerName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.ToString()}");
            }
            try
            {
                //Task.Delay(120000).Wait();
                obj.SentDataForReminderWACTA(triggerName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.ToString()}");
            }
        }
    }
}
