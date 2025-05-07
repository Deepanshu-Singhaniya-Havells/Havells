using System;
using System.ServiceModel.Description;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System.Configuration;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Text;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Data;
using System.Threading.Tasks;

namespace ConsoleApp
{
    class Program
    {
        #region Global Varialble declaration
        static IOrganizationService _service;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        #endregion  
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            _service = HavellsConnection.CreateConnection.createConnection(finalString);

            WhatsAppPromotions obj = new WhatsAppPromotions(_service);
            
            //obj.retriveJobs("5D_WP_AMC", -5, "72981d83-16fa-e811-a94c-000d3af0694e");
            //obj.retriveJobs("21D_WP_AMC", -21, "72981d83-16fa-e811-a94c-000d3af0694e");
            //obj.retriveJobs("5D_AC_AMC", -5, "D51EDD9D-16FA-E811-A94C-000D3AF0694E");
            obj.retriveJobs("21D_AC_AMC", -21, "D51EDD9D-16FA-E811-A94C-000D3AF0694E");

            Console.WriteLine("Success");
            Console.ReadLine();
        }

        #region Integration Configuration
        public static Integration IntegrationConfiguration(IOrganizationService _service, string Param)
        {
            Integration output = new Integration();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = _service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                if (integrationConfiguration.Contains("hil_username") && integrationConfiguration.Contains("hil_password"))
                {
                    output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }

        #endregion
    }
    public class Integration
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }


}

