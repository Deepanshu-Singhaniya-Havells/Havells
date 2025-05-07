using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.WebSockets;

namespace Havells.CRM.DynamicsPOSyncToSAP
{
    internal class Program
    {
        static IOrganizationService _service;
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        static void Main(string[] args)
        {
            _service = HelperClass.ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            //SyncPO_D365_to_SAP.GetProductRequisition(_service);
            Sync_Invoice_SAP_to_D365.GetAllInvoice(_service);
        }


    }

}
