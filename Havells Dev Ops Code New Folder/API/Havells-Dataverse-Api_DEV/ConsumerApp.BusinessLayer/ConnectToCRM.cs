using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;
using Microsoft.Xrm.Tooling.Connector;

namespace ConsumerApp.BusinessLayer
{
    public class ConnectToCRM
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        public static IOrganizationService GetOrgService()
        {
            var CrmURL = "https://havellscrmdev1.crm8.dynamics.com";
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = null;

            try
            {
                service = HavellsConnection.CreateConnection.createConnection(finalString);
            }
            catch
            {

            }
            return service;
        }

        public static IOrganizationService GetCRMConnection()
        {
            var CrmURL = "https://havellscrmdev1.crm8.dynamics.com";
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = null;

            try
            {
                service = HavellsConnection.CreateConnection.createConnection(finalString);
            }
            catch (Exception ex)
            {

            }
            return service;
        }

        public static IOrganizationService GetOrgServiceQA()
        {
            var CrmURL = "https://havellscrmdev1.crm8.dynamics.com";
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = null;

            try
            {
                service = HavellsConnection.CreateConnection.createConnection(finalString);
            }
            catch (Exception ex)
            {

            }
            return service;
        }

        public static IOrganizationService GetOrgServiceRebuild()
        {
            var CrmURL = "https://havellscrmdev1.crm8.dynamics.com";
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = null;

            try
            {
                service = HavellsConnection.CreateConnection.createConnection(finalString);
            }
            catch (Exception ex)
            {

            }
            return service;
        }
        public static IOrganizationService GetOrgServiceDev1()
        {
            var CrmURL = "https://havellscrmdev1.crm8.dynamics.com";
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = null;

            try
            {
                service = HavellsConnection.CreateConnection.createConnection(finalString);
            }
            catch (Exception ex)
            {

            }
            return service;
        }
        public static IOrganizationService GetOrgServiceProd()
        {
            var CrmURL = "https://havells.crm8.dynamics.com";
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = null;

            try
            {
                service = HavellsConnection.CreateConnection.createConnection(finalString);
            }
            catch (Exception ex)
            {

            }
            return service;
        }
    }

    public class APICredentials
    {
        public string APIURL { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
    }
}
