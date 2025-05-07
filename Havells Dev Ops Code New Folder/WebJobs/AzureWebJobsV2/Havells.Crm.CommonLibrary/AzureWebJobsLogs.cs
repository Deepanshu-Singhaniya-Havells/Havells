using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Auth;
using Microsoft.WindowsAzure.Storage.Blob;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace Havells.Crm.CommonLibrary
{
    public class AzureWebJobsLogs
    {
        public static string CreateOrUpdateLogs(string containerName, string fileName, string msg, string header)
        {
            #region AzureConnection
            string accName = "d365storagesa"; //This is your Azure Blob Storage Name
            string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
            string _recID = string.Empty;
            CloudStorageAccount strAcc = new CloudStorageAccount(new StorageCredentials(accName, accKey), true);
            CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
            CloudBlobContainer container = blobClient.GetContainerReference(containerName.ToLower());
            container.CreateIfNotExists();
            #endregion AzureConnection
            CloudAppendBlob appBlob = container.GetAppendBlobReference(fileName);
            if (!appBlob.Exists())
            {
                appBlob.CreateOrReplace();
                appBlob.AppendText
                (
                    string.Format("{0}\r\n", header)
                 );
            }
            appBlob.AppendText
            (
                string.Format("{0}\r\n", msg)
             );
            return appBlob.Uri.ToString();
        }
    }
}
