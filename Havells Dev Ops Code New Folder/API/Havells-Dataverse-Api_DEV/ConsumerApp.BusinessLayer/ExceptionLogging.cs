using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Auth;
using Microsoft.WindowsAzure.Storage.Blob;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
   public static class ExceptionLogging
    {
        private static String ErrorlineNo, Errormsg, extype, exurl, hostIp, ErrorLocation, HostAdd;

        public static void SendErrorToText(string msg, Exception ex = null)
        {
            #region AzureConnection
            string accName = "d365storagesa";
            string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
            string _recID = string.Empty;
            StorageCredentials creds = new StorageCredentials(accName, accKey);
            CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
            CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
            CloudBlobContainer container = blobClient.GetContainerReference("d365errorlog");
            container.CreateIfNotExistsAsync();
            #endregion AzureConnection

            var line = Environment.NewLine + Environment.NewLine;
            if (ex == null)
            {
                ErrorlineNo = "";
                Errormsg = "";
                extype = "";
                exurl = "";
                ErrorLocation = "";
            }
            else
            {
                ErrorlineNo = ex.StackTrace.Substring(ex.StackTrace.Length - 7, 7);
                Errormsg = ex.GetType().Name.ToString();
                extype = ex.GetType().ToString();
                exurl = "";
                ErrorLocation = ex.Message.ToString();
            }
            try
            {

                string rootDirDate = DateTime.Today.ToString("ddMMMyyyy");
                string filepathRootDate = @"C:/ErrorLog/" + rootDirDate;

                if (!Directory.Exists(filepathRootDate))
                {
                    Directory.CreateDirectory(filepathRootDate);
                }
                string rootDirDay = DateTime.Today.Day.ToString();
                string filepathRootDay = @"C:/ErrorLog/" + rootDirDate + "/" + rootDirDay;
                if (!Directory.Exists(filepathRootDay))
                {
                    Directory.CreateDirectory(filepathRootDay);
                }
                string rootDirHour = DateTime.Now.Hour.ToString();
                string filepathRootMainHour = @"C:/ErrorLog/" + rootDirDate + "/" + rootDirDay + "/" + rootDirHour;
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
                    if (ex == null)
                    {
                        string executionMsg = "Log Written Date:" + " " + DateTime.Now.ToString() + line + "Execution Message:" + " " + msg + line;
                        sw.WriteLine("-----------Executiion Message Details on " + " " + DateTime.Now.ToString() + "-----------------");
                        sw.WriteLine("-------------------------------------------------------------------------------------");
                        sw.WriteLine(line);
                        sw.WriteLine(executionMsg);
                        sw.WriteLine("--------------------------------*End*------------------------------------------");
                        sw.WriteLine(line);
                        sw.Flush();
                        sw.Close();

                        string path = rootDirDate + "/" + rootDirDay + "/" + rootDirHour + "/";
                        CloudAppendBlob appBlob = container.GetAppendBlobReference(path + DateTime.Today.ToString("dd-MM-yy") + ".txt");
                        if (!appBlob.Exists())
                        {
                            appBlob.CreateOrReplace();
                        }
                        string _strContent = sw.ToString();
                        _strContent = _strContent.Replace("@", System.Environment.NewLine);
                        appBlob.AppendText
                        (
                        string.Format(
                                "{0}\r\n",
                                _strContent)
                         );

                    }
                    else
                    {
                        string error = "Log Written Date:" + " " + DateTime.Now.ToString() + line + "Error Line No :" + " " + ErrorlineNo + line + "Error Message:" + " " + Errormsg + line + "Exception Type:" + " " + extype + line + "Error Location :" + " " + ErrorLocation + line + " Error Page Url:" + " " + exurl + line + "User Host IP:" + " " + hostIp + line;
                        sw.WriteLine("-----------Exception Details on " + " " + DateTime.Now.ToString() + "-----------------");
                        sw.WriteLine("-------------------------------------------------------------------------------------");
                        sw.WriteLine(line);
                        sw.WriteLine(error);
                        sw.WriteLine("--------------------------------*End*------------------------------------------");
                        sw.WriteLine(line);
                        sw.Flush();
                        sw.Close();

                        string path = rootDirDate + "/" + rootDirDay + "/" + rootDirHour + "/";
                        CloudAppendBlob appBlob = container.GetAppendBlobReference(path + DateTime.Today.ToString("dd-MM-yy") + ".txt");
                        if (!appBlob.Exists())
                        {
                            appBlob.CreateOrReplace();
                        }
                        string _strContent = sw.ToString();
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
            catch (Exception e)
            {
                e.ToString();

            }
        }
    }
}
