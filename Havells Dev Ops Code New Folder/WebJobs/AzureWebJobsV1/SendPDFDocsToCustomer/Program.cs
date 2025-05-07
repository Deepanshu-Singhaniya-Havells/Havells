using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Configuration;
using System.Net;
using System.Text;
using System.ServiceModel.Description;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace SendPDFDocsToCustomer
{
    public class Program
    {
        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            LoadAppSettings();
            if (loginUserGuid != Guid.Empty)
            {
                ConsumerSharables obj = new ConsumerSharables();
                obj.SendJobSheet(_service);
            }
        }

        static void SendJobSheetToConsumer() { 
        
        
        }
        #region App Setting Load/CRM Connection
        static void LoadAppSettings()
        {
            try
            {
                _userId = ConfigurationManager.AppSettings["CrmUserId"].ToString();
                _password = ConfigurationManager.AppSettings["CrmUserPassword"].ToString();
                _soapOrganizationServiceUri = ConfigurationManager.AppSettings["CrmSoapOrganizationServiceUri"].ToString();
                ConnectToCRM();
            }
            catch (Exception ex)
            {
                Console.WriteLine("SendPDFDocsToCustomer.Program.Main.LoadAppSettings ::  Error While Loading App Settings:" + ex.Message.ToString());
            }
        }
        static void ConnectToCRM()
        {
            try
            {
                ClientCredentials credentials = new ClientCredentials();
                credentials.UserName.UserName = _userId;
                credentials.UserName.Password = _password;
                Uri serviceUri = new Uri(_soapOrganizationServiceUri);
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);
                proxy.EnableProxyTypes();
                _service = (IOrganizationService)proxy;
                loginUserGuid = ((WhoAmIResponse)_service.Execute(new WhoAmIRequest())).UserId;
            }
            catch (Exception ex)
            {
                Console.WriteLine("SendPDFDocsToCustomer.Program.Main.ConnectToCRM :: Error While Creating Connection with MS CRM Organisation:" + ex.Message.ToString());
            }
        }
        #endregion
    }
}
