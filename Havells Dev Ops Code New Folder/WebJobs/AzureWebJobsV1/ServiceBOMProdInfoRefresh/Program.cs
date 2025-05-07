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

namespace ServiceBOMProdInfoRefresh
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
                RefreshServiceBOM();
            }
        }

        static void RefreshServiceBOM() {
            int totalRecords = 0;
            int recordCnt = 0;
            try
            {
                //string strFetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //<entity name='hil_servicebom'>
                //    <attribute name='hil_name' />
                //    <attribute name='createdon' />
                //    <attribute name='hil_product' />
                //    <attribute name='hil_servicebomid' />
                //    <attribute name='hil_sparepartfamily' />
                //    <attribute name='hil_productsubcategory' />
                //    <attribute name='hil_productcategory' />
                //<order attribute='createdon' descending='true' />
                //<filter type='and'>
                //    <condition attribute='hil_productsubcategory' operator='null' />
                //</filter>
                //<link-entity name='product' from='productid' to='hil_productcategory' visible='false' link-type='outer' alias='prd'>
                //    <attribute name='hil_materialgroup' />
                //</link-entity>
                //</entity>
                //</fetch>";
                string strFetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='hil_servicebom'>
                    <attribute name='hil_name' />
                    <attribute name='createdon' />
                    <attribute name='hil_sparepartfamily' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                      <condition attribute='hil_sparepartfamily' operator='null' />
                    </filter>
                    <link-entity name='product' from='productid' to='hil_product' link-type='inner' alias='prd'>
                      <attribute name='hil_sparepartfamily' />
                      <filter type='and'>
                        <condition attribute='hil_sparepartfamily' operator='not-null' />
                      </filter>
                    </link-entity>
                  </entity>
                </fetch>";
                while (true)
                {
                    Console.WriteLine("Job Started !!!");
                    EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(strFetchXML));
                    if (ec.Entities.Count > 0)
                    {
                        totalRecords = ec.Entities.Count;
                        recordCnt = 1;
                        foreach (Entity ent in ec.Entities)
                        {
                            //if (ent.Attributes.Contains("prd.hil_materialgroup"))
                            //{
                            //    ent["hil_productsubcategory"] = (EntityReference)(ent.GetAttributeValue<AliasedValue>("prd.hil_materialgroup").Value);
                            //    _service.Update(ent);
                            //}
                            //Console.WriteLine(ent.GetAttributeValue<string>("hil_name").ToString() + " : " + recordCnt.ToString() + "/" + totalRecords.ToString());
                            //recordCnt += 1;
                            if (ent.Attributes.Contains("prd.hil_sparepartfamily"))
                            {
                                ent["hil_sparepartfamily"] = (EntityReference)(ent.GetAttributeValue<AliasedValue>("prd.hil_sparepartfamily").Value);
                                _service.Update(ent);
                            }
                            Console.WriteLine(ent.GetAttributeValue<string>("hil_name").ToString() + " : " + recordCnt.ToString() + "/" + totalRecords.ToString());
                            recordCnt += 1;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Job Completed !!! No record found to sync");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error While Loading App Settings:" + ex.Message.ToString());
            }
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
                Console.WriteLine("ServiceBOMProdInfoRefresh.Program.Main.LoadAppSettings ::  Error While Loading App Settings:" + ex.Message.ToString());
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
                Console.WriteLine("ServiceBOMProdInfoRefresh.Program.Main.ConnectToCRM :: Error While Creating Connection with MS CRM Organisation:" + ex.Message.ToString());
            }
        }
        #endregion
    }
}
