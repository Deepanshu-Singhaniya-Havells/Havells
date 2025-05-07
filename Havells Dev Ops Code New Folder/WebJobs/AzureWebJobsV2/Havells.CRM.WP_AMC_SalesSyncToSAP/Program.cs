using Microsoft.Xrm.Sdk.Query;
using System.Text;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using System.Net;
using Microsoft.Crm.Sdk.Messages;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using System.Web.Services.Description;
using System.Configuration;
using System;
using Havells.Crm.CommonLibrary;
using System.Runtime.Remoting.Messaging;

namespace Havells.CRM.WP_AMC_SalesSyncToSAP
{
    public class Program : AzureWebJobsLogs
    {
        public static IOrganizationService service = null;
        static string containerName = "webjob-amcsapinvoices";
        public static string fileName = "";
        static string msg = null;
        static string header = null;
        static string blobUrl = null;
        static Program()
        {
            string path = DateTime.Now.ToString("ddMMMyyyy") + "/";
            fileName = path + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + ".csv";
            header = "Web Job Started";
            blobUrl = CreateOrUpdateLogs(containerName, fileName, msg, header);
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            service = Functions.createConnection(finalString, fileName);
        }
        static void Main(string[] args)
        {
            //Functions.getPaymentStatusJob(service, fileName);
            //Functions.getPaymentStatusofInvoice(service, fileName);
            Functions.PushAMCData_OmniChannel(service, fileName);
            //Functions.PushAMCData(service, fileName);
            //Functions.PushAMCData_WP(service, fileName);
        }

    }
}
