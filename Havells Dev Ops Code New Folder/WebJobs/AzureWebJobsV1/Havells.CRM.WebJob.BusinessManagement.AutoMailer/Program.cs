using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

namespace Havells.CRM.WebJob.BusinessManagement.AutoMailer
{
    public class Program
    {
        #region Global Varialble declaration
        static IOrganizationService service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion

        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        static void Main(string[] args)
        {
            Console.WriteLine("Program Started.");

            service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));

            Console.WriteLine("Connection is Established..");
            try
            {
                if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)service).IsReady)
                {
                    BOEReminder.SendBOEReminder(service);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("ERROR in Send BOE Reminder ***Error: " + ex.Message);
            }
        }
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
    }
}
