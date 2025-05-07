using Azure.Storage.Blobs;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.AuditLogMigration
{
    public static class Common
    {
        /// <summary>
        /// Validates the connection string information in app.config and throws an exception if it looks like 
        /// the user hasn't updated this to valid values. 
        /// </summary>
        /// <returns>BlobServiceClient object</returns>
        public static BlobServiceClient CreateblobServiceClientFromConnectionString()
        {
            BlobServiceClient blobServiceClient;
            const string Message = "Invalid storage account information provided. Please confirm the AccountName and AccountKey are valid in the app.config file.";

            try
            {
                blobServiceClient = new BlobServiceClient(ConfigurationManager.AppSettings.Get("StorageConnectionString"));
            }
            catch (FormatException)
            {
                Console.WriteLine(Message);
                Console.ReadLine();
                throw;
            }
            catch (ArgumentException)
            {
                Console.WriteLine(Message);
                Console.ReadLine();
                throw;
            }

            return blobServiceClient;
        }
    }
}
