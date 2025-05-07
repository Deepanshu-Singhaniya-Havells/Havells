using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using System.IO;
//using Microsoft.WindowsAzure.Storage.Blob;
//using Microsoft.WindowsAzure.Storage;
//using Microsoft.WindowsAzure.Storage.Auth;

namespace OMSMail
{
    public static class ExceptionLogging
    {

        private static String ErrorlineNo, Errormsg, extype, exurl, hostIp, ErrorLocation, HostAdd;

        public static void SendErrorToText(string msg)
        {
            //#region AzureConnection
            //string accName = "d365storagesa";
            //string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
            //string _recID = string.Empty;
            //StorageCredentials creds = new StorageCredentials(accName, accKey);
            //CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
            //CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
            //CloudBlobContainer container = blobClient.GetContainerReference("d365errorlog");
            //container.CreateIfNotExistsAsync();
            //#endregion AzureConnection

            var line = Environment.NewLine + Environment.NewLine;
           
                ErrorlineNo = "";
                Errormsg = "";
                extype = "";
                exurl = "";
                ErrorLocation = "";
            
            try
            {

                string rootDirDate = DateTime.Today.ToString("ddMMMyyyy");
                string filepathRootDate = @"C:/ErrorLog/" + rootDirDate ;  

                if (!Directory.Exists(filepathRootDate))
                {
                    Directory.CreateDirectory(filepathRootDate);
                }
                //string rootDirDay = DateTime.Today.Day.ToString();
                //string filepathRootDay = @"C:/ErrorLog/" + rootDirDate;// + "/" + rootDirDay;  
                //if (!Directory.Exists(filepathRootDay))
                //{
                //    Directory.CreateDirectory(filepathRootDay);
                //}
                //string rootDirHour =  DateTime.Now.Hour.ToString();
                string filepathRootMainHour = @"C:/ErrorLog/" + rootDirDate;// + "/" + rootDirDay+"/"+ rootDirHour;
                if (!Directory.Exists(filepathRootMainHour))
                {
                    Directory.CreateDirectory(filepathRootMainHour);
                }
                filepathRootMainHour = filepathRootMainHour + "/" + DateTime.Today.ToString("dd-MM-yy") + ".txt";  
                if (!File.Exists(filepathRootMainHour))
                {
                    File.Create(filepathRootMainHour).Dispose();
                }
                
                using (StreamWriter sw = File.AppendText(filepathRootMainHour))
                {
                    //if (ex == null)
                    //{
                        //string executionMsg = "Log Written Date:" + " " + DateTime.Now.ToString() + line + "Execution Message:" + " " + msg + line;
                        //sw.WriteLine("-----------Executiion Message Details on " + " " + DateTime.Now.ToString() + "-----------------");
                        //sw.WriteLine("-------------------------------------------------------------------------------------");
                        //sw.WriteLine(line);
                        sw.WriteLine(msg);
                        //sw.WriteLine("--------------------------------*End*------------------------------------------");
                        //sw.WriteLine(line);
                        sw.Flush();
                        sw.Close();

                        //string path = rootDirDate + "/" + rootDirDay + "/" + rootDirHour+"/";
                        //CloudAppendBlob appBlob = container.GetAppendBlobReference(path+ DateTime.Today.ToString("dd-MM-yy") + ".txt");
                        //if (!appBlob.Exists())
                        //{
                        //    appBlob.CreateOrReplace();
                        //}
                        //string _strContent = sw.ToString();
                        //_strContent = _strContent.Replace("@", System.Environment.NewLine);
                        //appBlob.AppendText
                        //(
                        //string.Format(
                        //        "{0}\r\n",
                        //        _strContent)
                        // );

                    //}
                    //else
                    //{
                    //    string error = "Log Written Date:" + " " + DateTime.Now.ToString() + line + "Error Line No :" + " " + ErrorlineNo + line + "Error Message:" + " " + Errormsg + line + "Exception Type:" + " " + extype + line + "Error Location :" + " " + ErrorLocation + line + " Error Page Url:" + " " + exurl + line + "User Host IP:" + " " + hostIp + line;
                    //    sw.WriteLine("-----------Exception Details on " + " " + DateTime.Now.ToString() + "-----------------");
                    //    sw.WriteLine("-------------------------------------------------------------------------------------");
                    //    sw.WriteLine(line);
                    //    sw.WriteLine(error);
                    //    sw.WriteLine("--------------------------------*End*------------------------------------------");
                    //    sw.WriteLine(line);
                    //    sw.Flush();
                    //    sw.Close();

                    //    string path = rootDirDate + "/" + rootDirDay + "/" + rootDirHour + "/";
                    //    //CloudAppendBlob appBlob = container.GetAppendBlobReference(path + DateTime.Today.ToString("dd-MM-yy") + ".txt");
                    //    //if (!appBlob.Exists())
                    //    //{
                    //    //    appBlob.CreateOrReplace();
                    //    //}
                    //    //string _strContent = sw.ToString();
                    //    //_strContent = _strContent.Replace("@", System.Environment.NewLine);
                    //    //appBlob.AppendText
                    //    //(
                    //    //string.Format(
                    //    //        "{0}\r\n",
                    //    //        _strContent)
                    //    // );
                    //}
                }

            }
            catch (Exception e)
            {
                e.ToString();

            }
        }
        //private static CloudBlobContainer ConnectWithAzureBlob()
        //{
        //    string ConnectionSting = "DefaultEndpointsProtocol=https;AccountName=d365storagesa;AccountKey=6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==;EndpointSuffix=core.windows.net";
        //    CloudAppendBlob appendBlobReference = null;
        //    try
        //    {
        //        CloudStorageAccount storageaccount = CloudStorageAccount.Parse(ConnectionSting);
        //        CloudBlobClient client = storageaccount.CreateCloudBlobClient();
        //        //Get the reference of the Container. The GetConainerReference doesn't make a request to the Blob Storage but the Create() & CreateIfNotExists() method does. The method CreateIfNotExists() could be use whether the Container exists or not  
        //        CloudBlobContainer containerRoot = client.GetContainerReference("errorlog");
        //        containerRoot.CreateIfNotExists();
        //        containerRoot.SetPermissions(new BlobContainerPermissions { PublicAccess = BlobContainerPublicAccessType.Container });
        //        Console.WriteLine("Container got created successfully");

        //    }
        //    catch { }
        //    return containerRoot;
        //}
    }
}
