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


namespace AssignmentMatrixUpcountryUpdate
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
                UpdateAssignmentMatrixUpcountry();
            }
        }

        static void UpdateAssignmentMatrixUpcountry()
        {
            int totalRecords = 0;
            int recordCnt = 0;
            int recordCnt1 = 0;
            try
            {
                string strFetchXML1 = string.Empty;
                string strPINCode = string.Empty;
                string strASPCode = string.Empty;
                string strRemarks = string.Empty;

                string strFetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='hil_upcountrydataupdate'>
                    <attribute name='hil_upcountrydataupdateid' />
                    <attribute name='hil_name' />
                    <attribute name='createdon' />
                    <attribute name='hil_remarks' />
                    <attribute name='hil_channelpartnercode' />
                    <attribute name='hil_brand' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='statecode' operator='eq' value='0' />
                        <condition attribute='hil_remarks' operator='null'/>
                    </filter>
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
                            if (ent.Attributes.Contains("hil_name"))
                            {
                                strPINCode = (ent.GetAttributeValue<string>("hil_name").ToString());
                            }
                            if (ent.Attributes.Contains("hil_channelpartnercode"))
                            {
                                strASPCode = (ent.GetAttributeValue<string>("hil_channelpartnercode").ToString());
                            }

                            strFetchXML1 = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_assignmentmatrix'>
                                <attribute name='hil_assignmentmatrixid' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                  <condition attribute='hil_pincodename' operator='like' value='%" + strPINCode + @"%' />
                                  <condition attribute='hil_franchiseedirectengineername' operator='like' value='%" + strASPCode + @"%' />
                                  <condition attribute='statecode' operator='eq' value='0' />
                                </filter>
                                <link-entity name='product' from='productid' to='hil_division' link-type='outer' alias='ab'>
                                <filter type='and'>
                                    <condition attribute='hil_brandidentifier' operator='eq' value='2' />
                                </filter>
                                </link-entity>                             
                             </entity>
                            </fetch>";
                            Console.WriteLine("Inner Job Started !!!");
                            EntityCollection ec1 = _service.RetrieveMultiple(new FetchExpression(strFetchXML1));
                            if (ec1.Entities.Count > 0)
                            {
                                recordCnt1 = 1;
                                foreach (Entity ent1 in ec1.Entities)
                                {
                                    ent1["hil_upcountry"] = true;
                                    _service.Update(ent1);
                                    recordCnt1 += 1;
                                    Console.WriteLine(recordCnt1.ToString() + "/" + recordCnt.ToString() + "/" + totalRecords.ToString());
                                }
                                // Set Record status : Inactive
                                SetStateRequest setStateRequest = new SetStateRequest()
                                {
                                    EntityMoniker = new EntityReference
                                    {
                                        Id = ent.Id,
                                        LogicalName = "hil_upcountrydataupdate",
                                    },
                                    State = new OptionSetValue(1), //Inactive
                                    Status = new OptionSetValue(2) //Inactive
                                };
                                _service.Execute(setStateRequest);
                            }
                            else {
                                ent["hil_remarks"] = "Invalid PIN Code/ASP Code";
                                _service.Update(ent);
                            }
                            Console.WriteLine(recordCnt.ToString() + "/" + totalRecords.ToString());
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
                Console.WriteLine("AssignmentMatrixUpcountryUpdate.Program.Main.LoadAppSettings ::  Error While Loading App Settings:" + ex.Message.ToString());
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
                Console.WriteLine("AssignmentMatrixUpcountryUpdate.Program.Main.ConnectToCRM :: Error While Creating Connection with MS CRM Organisation:" + ex.Message.ToString());
            }
        }
        #endregion
    }
}
