
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.WebServiceClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Crm.Sdk.Messages;
using System;

namespace Havells.CRM.EnviromentMetaDateMigration
{
    internal class Program
    {
        public const int _languageCode = 1033;
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        public static string preFix = null;
        public static Guid PublisherSer = new Guid("D21AAB71-79E7-11DD-8874-00188B01E34F");
        public static Guid PublisherPrd = new Guid("D21AAB71-79E7-11DD-8874-00188B01E34F");
        public static EntityReference SystemAdmin = new EntityReference("systemuser", new Guid("1a8fc0e8-7e48-ed11-bba2-6045bdac5a88"));

        static void Main(string[] args)
        {
            IOrganizationService _SourceService = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            IOrganizationService _DestinationService = ConnectToCRM(string.Format(connStr, "https://havellsservice.crm8.dynamics.com"));
            // GlobleOptionSet.CreateGlobalOptionSet(_SourceService, _DestinationService);
            //WebresourceMigration.UpsertWebResource(_SourceService, _DestinationService);
            Entities.EntityMigration(_SourceService, _DestinationService);

        }
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            //  GetCRMService();
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);

                //var sdkService = new OrganizationWebProxyClient(GetServiceUrl(organizationUrl), new TimeSpan(0, 10, 0), false);


                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    WriteLogFile.WriteLog(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                WriteLogFile.WriteLog("-------------------------Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
    }
}
