using Microsoft.Xrm.Sdk;
using System;
using System.Configuration;

namespace Havells.CRM.WebJob.MasterSync
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
                var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
                string finalString = string.Format(connStr, CrmURL);
                IOrganizationService service = HavellsConnection.CreateConnection.createConnection(finalString);
                //State.synsState(service);
                //City.syncCity(service);
                //District.syncDistrict(service);
                //Pincode.syncPinCode(service);
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error " + ex.Message);
            }
           
        }
    }
}
