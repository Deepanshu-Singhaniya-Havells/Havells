using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System.Configuration;

namespace Havells.CRM.ProductAndPriceListSync
{
    class Program
    {
        
        public static IOrganizationService _service;

        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();//"AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8="
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();//"https://havells.crm8.dynamics.com"
            var cableDivisionGUID = ConfigurationManager.AppSettings["cableDivisionGUID"].ToString();//"fd555381-16fa-e811-a94d-000d3af06cd4";
            var motorDivisionGuid = ConfigurationManager.AppSettings["motorDivisionGuid"].ToString();//"E0E38C8E-16FA-E811-A94C-000D3AF06C56";
            _service = Models.ConnectToCRM(string.Format(connStr, CrmURL));
            if (((CrmServiceClient)_service).IsReady)
            {
                SyncProducts.syncProducts(_service, (args.Length > 0 ? args[0] : ""));
                //PriceListSync.SparePartPriceList(_service, (args.Length > 0 ? args[0] : ""), motorDivisionGuid, cableDivisionGUID);//motor Division Guid 
                //PriceListSync.AMCPriceList(_service);
            }
        }
    }
}
