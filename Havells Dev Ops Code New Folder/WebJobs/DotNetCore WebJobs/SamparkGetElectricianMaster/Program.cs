using Microsoft.PowerPlatform.Dataverse.Client;
using System.Configuration;

namespace SamparkGetElectricianMaster
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"];
            var CRMUrl = ConfigurationManager.AppSettings["CRMUrl"];
            string finalConnStr = string.Format(connStr ?? "", CRMUrl);
            ServiceClient service = new ServiceClient(finalConnStr);
            ClsElectricianMasterSampark obj = new ClsElectricianMasterSampark(service);
            obj.GetElectricianMasterForSampark(); //calling method
        }
    }
}
