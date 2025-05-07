using Microsoft.WindowsAzure.Storage.Auth;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.WindowsAzure.Storage;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportChannelPartnerSampark
{
    public class WriteLogFile
    {
        public static bool WriteLog(string strMessage, string strFileName)
        {
            Console.WriteLine(strMessage);
            return true;
            //try
            //{
            //    //string strFileName = "_InstanceMigration_Logs_" + DateTime.Today.ToString("yyyy/MM/dd") + ".txt";
            //    FileStream objFilestream = new FileStream(strFileName, FileMode.Append, FileAccess.Write);
            //    StreamWriter objStreamWriter = new StreamWriter((Stream)objFilestream);
            //    objStreamWriter.WriteLine(strMessage);
            //    objStreamWriter.Close();
            //    objFilestream.Close();
            //    return true;
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //    return false;
            //}
        }
        public static void WriteLogToBlob(string strContent, string fileName)
        {
            try
            {
                string accName = "d365storagesa";
                string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
                //bool _retValue = false;
                string _recID = string.Empty;

                // Implement the accout, set true for https for SSL.  
                StorageCredentials creds = new StorageCredentials(accName, accKey);
                CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
                // Create the blob client.  
                CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
                // Retrieve a reference to a container.   
                CloudBlobContainer container = blobClient.GetContainerReference("a-logs-cp");
                // Create the container if it doesn't already exist.  
                container.CreateIfNotExistsAsync();
                if (!CheckIfAzureBlobExist(fileName, container))
                {
                    UploadAzureBlob(fileName, strContent.ToString(), container);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Update Logs on Blob " + ex.Message);
            }
        }
        static bool CheckIfAzureBlobExist(string _fileName, CloudBlobContainer container)
        {
            bool _retValue = false;
            CloudAppendBlob appBlob = container.GetAppendBlobReference(_fileName + ".csv");

            if (appBlob.Exists())
            {
                _retValue = true;
            }
            return _retValue;
        }
        static void UploadAzureBlob(string _fileName, string _strContent, CloudBlobContainer container)
        {
            // This creates a reference to the append blob.  
            CloudAppendBlob appBlob = container.GetAppendBlobReference(_fileName);

            // Now we are going to check if todays file exists and if it doesn't we create it.  
            if (!appBlob.Exists())
            {
                appBlob.CreateOrReplace();
            }

            // Add the entry to file.  
            _strContent = _strContent.Replace("@", System.Environment.NewLine);
            appBlob.AppendText
            (
            string.Format(
                    "{0}\r\n",
                    _strContent)
             );
        }
    }
}
