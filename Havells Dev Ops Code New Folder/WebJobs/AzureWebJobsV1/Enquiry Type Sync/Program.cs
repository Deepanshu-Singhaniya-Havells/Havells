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
using Microsoft.Xrm.Sdk.Query;
using System.IO;
using Newtonsoft.Json;
using System.Runtime.Serialization.Json;
using Enquiry_Type_Sync.Model;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Enquiry_Type_Sync
{
    class Program
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
                SyncEnquiryTypeData();
            }
        }
        private static void SyncEnquiryTypeData()
        {
            WebClient webClient = new WebClient();
            string url = GetIntegrationUrl(_service, "EnquiryType");
            //string url = GetIntegrationUrl(_service, "EnquiryType");
            string RunDate;
            string _globalOptionSetName = "hil_leadtype";
            QueryExpression Query = new QueryExpression("hil_enquirytype");
            Query.ColumnSet = new ColumnSet("hil_enquirytypecode");
            EntityCollection Found = _service.RetrieveMultiple(Query);
            bool _publish = false;

            if (Found.Entities.Count == 0)
            {
                RunDate = "19000101000000";
            }
            else
            {
                RunDate = DateTime.Now.Date.AddDays(-1).ToString("yyyyMMdd") + "000000";
            }

            url = url + RunDate;

            //Implemented Basic Authentication
            string authInfo = "D365_HAVELLS" + ":" + "DEVD365@1234";
            authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
            webClient.Headers["Authorization"] = "Basic " + authInfo;

            var jsonData = webClient.DownloadData(url);
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(RootObject));
            RootObject rootObject = (RootObject)ser.ReadObject(new MemoryStream(jsonData));

            int iDone = 0;
            int iTotal = rootObject.Results.Count;
            int UpdateCount = 0;
            string enquiryType = string.Empty;

            Console.WriteLine("Total Count: " + iTotal.ToString());

            Guid IfExists = Guid.Empty;

            foreach (EnquiryTypeData obj in rootObject.Results)
            {
                enquiryType = obj.EnquiryTypeCode;
                IfExists = CheckIfEnquiryTypeExists(enquiryType, _service);
                if (IfExists == Guid.Empty) // New Creation
                {
                    Entity ent = new Entity("hil_enquirytype");
                    ent["hil_enquirytypecode"] = Int32.Parse(obj.EnquiryTypeCode) ;
                    ent["hil_enquirytypename"] = obj.EnquiryTypeDesc;
                    _service.Create(ent);
                    iDone += 1;
                    Console.WriteLine("New Record Count: " + iDone.ToString() + "/" + iTotal.ToString());
                }
                else {  // Updation
                    Entity ent = new Entity("hil_enquirytype");
                    ent.Id = IfExists;
                    ent["hil_enquirytypename"] = obj.EnquiryTypeDesc;
                    _service.Update(ent);
                    UpdateCount += 1;
                    Console.WriteLine("Update Record Count: " + UpdateCount.ToString() + "/" + iTotal.ToString());
                }

                #region Create Options in GlobalOptionSet
                RetrieveOptionSetRequest retrieveOptionSetRequest =
                new RetrieveOptionSetRequest
                {
                    Name = _globalOptionSetName
                };
                // Execute the request.
                RetrieveOptionSetResponse retrieveOptionSetResponse = (RetrieveOptionSetResponse)_service.Execute(retrieveOptionSetRequest);
                //Console.WriteLine("Retrieved {0}.", retrieveOptionSetRequest.Name);
                // Access the retrieved OptionSetMetadata.
                OptionSetMetadata retrievedOptionSetMetadata =
                (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;
                // Get the current options list for the retrieved attribute.
                OptionMetadata[] optionList =
                retrievedOptionSetMetadata.Options.ToArray();
                if (Array.Find(optionList, o => o.Value == Convert.ToInt32(obj.EnquiryTypeCode)) == null)
                {
                    InsertOptionValueRequest insertOptionValueRequest =
                    new InsertOptionValueRequest
                    {
                        OptionSetName = _globalOptionSetName,
                        Label = new Label(obj.EnquiryTypeDesc, 1033),
                        Value = Convert.ToInt32(obj.EnquiryTypeCode)
                    };
                    int _insertedOptionValue = ((InsertOptionValueResponse)_service.Execute(
                    insertOptionValueRequest)).NewOptionValue;
                    _publish = true;
                }
                else
                {
                    UpdateOptionValueRequest updateOptionValueRequest =
                    new UpdateOptionValueRequest
                    {
                        OptionSetName = _globalOptionSetName,
                        Value = Convert.ToInt32(obj.EnquiryTypeCode),
                        Label = new Label(obj.EnquiryTypeDesc, 1033)
                    };
                    _service.Execute(updateOptionValueRequest);
                    _publish = true;
                }
                if (_publish)
                {
                    //Publish the OptionSet
                    PublishXmlRequest pxReq2 = new PublishXmlRequest { ParameterXml = String.Format("<importexportxml><optionsets><optionset>{0}</optionset></optionsets></importexportxml>", _globalOptionSetName) };
                    _service.Execute(pxReq2);
                }
                #endregion
            }
            Console.WriteLine("Done");
        }

        public static Guid CheckIfEnquiryTypeExists(string enquirytypecode, IOrganizationService service)
        {
            Guid EnquiryType = Guid.Empty;
            QueryExpression Query = new QueryExpression("hil_enquirytype");
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_enquirytypecode", ConditionOperator.Equal, enquirytypecode);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                EnquiryType = Found.Entities[0].Id;
            }
            return EnquiryType;
        }

        #region GET INTEGRATION URL
        public static string GetIntegrationUrl(IOrganizationService service, string RecName)
        {
            string sUrl = string.Empty;
            QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_url");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, RecName);
            Query.Criteria.AddCondition("hil_url", ConditionOperator.NotNull);
            EntityCollection Found = service.RetrieveMultiple(Query);
            hil_integrationconfiguration enConfig = Found.Entities[0].ToEntity<hil_integrationconfiguration>();
            sUrl = enConfig.hil_Url;
            return sUrl;
        }
        #endregion

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
                Console.WriteLine("EnquiryType_Sync.Program.Main.LoadAppSettings ::  Error While Loading App Settings:" + ex.Message.ToString());
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
                Console.WriteLine("EnquiryType_Sync.Program.Main.ConnectToCRM :: Error While Creating Connection with MS CRM Organisation:" + ex.Message.ToString());
            }
        }
        #endregion

    }
}
