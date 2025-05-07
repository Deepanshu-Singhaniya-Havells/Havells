using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using System.Configuration;

namespace Havells.CRM.AuditLogMigration
{
    public class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Azure Blob Storage - Getting Started Samples\n");
            GettingStarted.CallBlobGettingStartedSamples();

            Console.WriteLine("Azure Blob Storage - Advanced Samples\n ");
            Advanced.CallBlobAdvancedSamples().Wait();

            Console.WriteLine();
            Console.WriteLine("Press any key to exit.");
            Console.ReadLine();

            //var connStr = ConfigurationManager.AppSettings["connStr"];
            //var CrmURLProd = ConfigurationManager.AppSettings["CRMUrl"];
            //string finalConnStr = string.Format(connStr ?? "", CrmURLProd);
            //ServiceClient service = new ServiceClient(finalConnStr);

            //AuditLogMigration obj = new AuditLogMigration(service);


            //var fromDate = ConfigurationManager.AppSettings["fromDate"].ToString();
            //var toDate = ConfigurationManager.AppSettings["toDate"].ToString();
            //var objecttypeCode = ConfigurationManager.AppSettings["ObjectTypeCode"].ToString();

          
            ////ListAllAuditEnableEntity(service);

            //////Delete Transactional Data
            ////DeleteTransactionalData(service, "msdyn_workorderservice", fromDate, toDate);

            ////GetAuditLogHeader(service, int.Parse(objecttypeCode), fromDate, toDate);
            //DateTime fD = DateTime.Today;
            //DateTime tD = DateTime.Today;
            //fD = fD.AddHours(Convert.ToInt32(ConfigurationManager.AppSettings["startHour"]));
            //tD = tD.AddHours(Convert.ToInt32(ConfigurationManager.AppSettings["endHour"]));
            //obj.GetAuditLogByEntityDailyLogs(int.Parse(objecttypeCode), fD.ToString("yyyy-MM-dd HH:mm:ss"), tD.ToString("yyyy-MM-dd HH:mm:ss"));

            //fD = fD.AddHours(6);
            //tD = tD.AddHours(6);
            //GetAuditLogByEntityDailyLogs(service, int.Parse(objecttypeCode), fD.ToString(), tD.ToString());

            //fD = fD.AddHours(6);
            //tD = tD.AddHours(6);
            //GetAuditLogByEntityDailyLogs(service, int.Parse(objecttypeCode), fD.ToString(), tD.ToString());

            //fD = fD.AddHours(6);
            //tD = tD.AddHours(6);
            //GetAuditLogByEntityDailyLogs(service, int.Parse(objecttypeCode), fD.ToString(), tD.ToString());

            //GetAuditLogByEntityDailyLogs(service, int.Parse(objecttypeCode), fromDate, toDate);

            //GetAllAuditEnableEntity(service, fromDate, toDate);

            //updateFlag(entityName, fromDate, toDate, service);
            //D365AuditLogMigration_DeltaWithFlag("msdyn_customerasset", "2021-07-01", "2021-07-31", service);

            //GetAuditEnableEntity(service);
            //foreach (KeyValuePair<string, int?> kvp in AuditLogEnabledEntityList)
            //{
            //    Console.WriteLine("Key = {0}, Value = {1}", kvp.Key, kvp.Value);
            //    //D365AuditLogMigration_Delta(kvp.Key, service);
            //}


        }
    }
    public class EntityAttribute
    {
        public string LogicalName { get; set; }
        public string DisplayName { get; set; }
        public string AttributeType { get; set; }
        public List<EntityOptionSet> Optionset { get; set; }
    }
    public class EntityOptionSet
    {
        public string OptionName { get; set; }
        public int? OptionValue { get; set; }
    }

}
