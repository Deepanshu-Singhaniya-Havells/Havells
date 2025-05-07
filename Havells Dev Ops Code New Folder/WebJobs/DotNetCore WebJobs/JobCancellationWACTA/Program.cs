using Microsoft.PowerPlatform.Dataverse.Client;
using System.Configuration;

namespace JobCancellationWACTA
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"];
            var CrmURL = ConfigurationManager.AppSettings["CRMUrlProd"];
            string finalConnStr = string.Format(connStr ?? "", CrmURL);
            string triggerNames = ConfigurationManager.AppSettings["triggerNames"];
            string triggerNameForDuplicate = ConfigurationManager.AppSettings["triggerNameForDuplicate"];
            ServiceClient service = new ServiceClient(finalConnStr);

            WATriggerOnJobCancelReason obj = new WATriggerOnJobCancelReason(service);

            // trigger on basis of Job Cancel reason
            try
            {
                foreach (string triggerName in triggerNames.Split(';'))
                {
                    obj.SentDataForRequestForCancellation(triggerName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.ToString()}");
            }

            // trigger on basis of Job Cancel reason

            //try
            //{
            //    foreach (string triggerName in triggerNames.Split(';'))
            //    {
            //        obj.SentDataForReminderWACTA(triggerName);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"Error: {ex.ToString()}");
            //}

            //  trigger when Job is Duplicate
            try
            {
                obj.SentDataForDuplicateRequest(triggerNameForDuplicate);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.ToString()}");
            }
        }
    }
}
