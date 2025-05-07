using System;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using System.ServiceModel.Description;
using System.Net;
using System.Linq;
using Microsoft.Xrm.Sdk.Query;
using System.Xml.Linq;
using Havells_Schedulers;

namespace Havells_Schedulers
{
    class Program
    {

        public static IOrganizationService ConnecttoCRM()
        {
            IOrganizationService organizationService = null;

            //String username = "pwcuser1@havells.com";//eg: abc@xyz.onmicrosoft.com
            //String password = "havells@123";//eg: password@123

            // Get the URL from CRM, Navigate to Settings -> Customizations -> Developer Resources
            // Copy and Paste Organization Service Endpoint Address URL
            String url = "https://havells.api.crm8.dynamics.com/XRMServices/2011/Organization.svc"; //eg: https://<yourorganisationname>.api.crm8.dynamics.com/XRMServices/2011/Organization.svc
            try
            {
                ClientCredentials clientCredentials = new ClientCredentials();
                clientCredentials.UserName.UserName = "pwcuser1@havells.com";
                clientCredentials.UserName.Password = "havells@123";

                // For Dynamics 365 Customer Engagement V9.X, set Security Protocol as TLS12
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                OrganizationServiceProxy orgProxy = new OrganizationServiceProxy(new Uri(url), null, clientCredentials, null);
                //serviceProxy.EnableProxyTypes();
                orgProxy.EnableProxyTypes();
                organizationService = (IOrganizationService)orgProxy;

                if (organizationService != null)
                {
                    Guid gOrgId = ((WhoAmIResponse)organizationService.Execute(new WhoAmIRequest())).OrganizationId;
                    if (gOrgId != Guid.Empty)
                    {
                        Console.WriteLine("Connection Established Successfully...");
                    }
                }
                else
                {
                    Console.WriteLine("Failed to Established Connection!!!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception occured - " + ex.Message);
            }
            return organizationService;

        }

        static void Main(string[] args)
        {
            IOrganizationService service = ConnecttoCRM();

            Entity entity = service.Retrieve(Contact.EntityLogicalName, new Guid("58585BE5-FB08-E911-A94D-000D3AF03089"), new ColumnSet(new string[] { "firstname" }));

            bool att = entity.Attributes.Contains("fullname");
            bool value = entity.Contains("fullname");
            //Havells_Schedulers.MSLExecution.InitiateExecution(service);
        }

    }
}