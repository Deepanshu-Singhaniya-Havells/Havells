using ChannelPartnerSyncJob.Models;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

namespace ChannelPartnerSyncJob
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
                GetChannelPartnerData(_service, (args.Length > 0 ? args[0] : ""));
            }
        }
        public static void GetChannelPartnerData(IOrganizationService service, string _syncDatetime)
        {
            try
            {
                Integration intConf = GetIntegration(service, "Tier 1 Customer");
                string uri = intConf.uri;
                string authInfo = intConf.userName + ":" + intConf.passWord;
                authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));

                if (_syncDatetime != string.Empty && _syncDatetime.Trim().Length > 0)
                {
                    uri = uri + _syncDatetime;
                }
                else
                {
                    uri = uri + getTimeStamp(service);
                }
                Console.WriteLine("URL: " + uri);

                var client = new RestClient(uri);
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);
                request.AddHeader("Authorization", "Basic " + authInfo);
                Console.WriteLine("Downloading Channel Partner Data: " + DateTime.Now.ToString());
                IRestResponse response = client.Execute(request);
                PartnerRootObject rootObject = JsonConvert.DeserializeObject<PartnerRootObject>(response.Content);
                Console.WriteLine("Downloading Completed of Channel Partner Data: " + DateTime.Now.ToString());
                SyncChannelPartner(rootObject, service);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
            }
        }

        private static void SyncChannelPartner(PartnerRootObject rootObject, IOrganizationService service)
        {
            string _KTOKD = string.Empty;
            Guid IfExists = Guid.Empty;
            string iKUNNR = string.Empty;
            Account iAccount = new Account();
            EntityReference entRefDistrict = null;
            int iDone = 0;
            int iTotal = rootObject.Results.Count;
            foreach (PartnerResult obj in rootObject.Results)
            {
                try
                {
                    _KTOKD = obj.KTOKD.ToString().Trim();
                    iKUNNR = obj.KUNNR;
                    Console.WriteLine("Channel Partner Code: " + iKUNNR);
                    if (iKUNNR.StartsWith("F"))
                        iKUNNR = iKUNNR.Substring(1);
                    IfExists = CheckIfPartnerExists(iKUNNR, service);
                    OptionSetValue _customerType = new OptionSetValue(_KTOKD == "0010" ? 1 : _KTOKD == "0020" ? 13 : _KTOKD == "0050" ? 14 : _KTOKD == "0045" ? 15 : _KTOKD == "0030" ? 16 : _KTOKD == "0031" ? 17 : _KTOKD == "0056" ? 6 : _KTOKD == "0065" ? 9 : 12);
                    if (_customerType.Value != 12)
                    {
                        if (obj.delete_flag != "X")
                        {
                            iAccount.hil_InWarrantyCustomerSAPCode = obj.KUNNR;
                            iAccount.hil_OutWarrantyCustomerSAPCode = iKUNNR;
                            iAccount.Name = obj.VTXTM;
                            iAccount.EMailAddress1 = obj.SMTP_ADDR;
                            iAccount.Telephone1 = obj.MOB_NUMBER;
                            iAccount.Address1_Line1 = obj.STREET;
                            iAccount.Address1_Line2 = obj.STR_SUPPL3;
                            iAccount.Address1_Line3 = obj.ADDRESS3;
                            iAccount.AccountNumber = iKUNNR;
                            iAccount.hil_StagingPinUniqueKey = obj.KTOKD;
                            if (obj.Mtimestamp == null)
                                iAccount["hil_mdmtimestamp"] = ConvertToDateTime(obj.Ctimestamp);
                            else
                                iAccount["hil_mdmtimestamp"] = ConvertToDateTime(obj.Mtimestamp);

                            iAccount["address1_postofficebox"] = obj.GST_NO;
                            iAccount["hil_pan"] = obj.J_1IPANNO;

                            QueryExpression Query = new QueryExpression("hil_businessmapping");
                            Query.ColumnSet = new ColumnSet(true);
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition(new ConditionExpression("hil_stagingarea", ConditionOperator.Equal, obj.DM_AREA));
                            Query.Criteria.AddCondition(new ConditionExpression("hil_stagingpin", ConditionOperator.Equal, obj.dm_pin));
                            EntityCollection Found = service.RetrieveMultiple(Query);
                            if (Found.Entities.Count > 0)
                            {
                                hil_businessmapping iBusMap = Found.Entities[0].ToEntity<hil_businessmapping>();
                                iAccount.hil_city = iBusMap.hil_city;
                                iAccount.hil_area = iBusMap.hil_area;
                                iAccount.hil_pincode = iBusMap.hil_pincode;
                                iAccount.hil_region = iBusMap.hil_region;
                                iAccount.hil_state = iBusMap.hil_state;
                                entRefDistrict = iBusMap.hil_district;
                                iAccount.hil_branch = iBusMap.hil_branch;
                                iAccount.hil_subterritory = iBusMap.hil_subterritory;
                                iAccount.hil_salesoffice = iBusMap.hil_salesoffice;
                            }
                            if (obj.Longitude != null)
                                iAccount.Address1_Longitude = Double.Parse(obj.Longitude);
                            if (obj.Latitude != null)
                                iAccount.Address1_Latitude = Double.Parse(obj.Latitude);

                            if (IfExists == Guid.Empty)
                            {
                                iAccount.hil_district = entRefDistrict;
                                iAccount.CustomerTypeCode = _customerType;// new OptionSetValue(_KTOKD == "0010" ? 1 : _KTOKD == "0020" ? 13 : _KTOKD == "0050" ? 14 : _KTOKD == "0045" ? 15 : _KTOKD == "0030" ? 16 : _KTOKD == "0031" ? 17 : _KTOKD == "0056" ? 6 : 9);
                                service.Create(iAccount);
                            }
                            else
                            {
                                //iAccount.Id = IfExists;
                                //service.Update(iAccount);
                            }
                        }
                        else if (obj.delete_flag == "X")
                        {
                            if (IfExists != Guid.Empty)
                            {
                                SetStateRequest setStateRequest = new SetStateRequest()
                                {
                                    EntityMoniker = new EntityReference
                                    {
                                        Id = IfExists,
                                        LogicalName = "account"
                                    },
                                    State = new OptionSetValue(1), //deactive
                                    Status = new OptionSetValue(2) //deactive
                                };
                                service.Execute(setStateRequest);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error :" + ex.Message);
                }
                iDone = iDone + 1;
                Console.WriteLine("Record has been processed :" + iDone + "/" + iTotal);
            }
            Console.WriteLine(" TOTAL COUNT :" + iTotal.ToString());
        }
        public static DateTime? ConvertToDateTime(string _mdmTimeStamp)
        {
            DateTime? _dtMDMTimeStamp = null;
            try
            {
                _dtMDMTimeStamp = new DateTime(Convert.ToInt32(_mdmTimeStamp.Substring(0, 4)), Convert.ToInt32(_mdmTimeStamp.Substring(4, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(6, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(8, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(10, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(12, 2)));
            }
            catch { }
            return _dtMDMTimeStamp;
        }
        public static string getTimeStamp(IOrganizationService service)
        {
            string _enquiryDatetime = "20210804000000";
            QueryExpression qsCType = new QueryExpression("account");
            qsCType.ColumnSet = new ColumnSet("hil_mdmtimestamp");
            qsCType.NoLock = true;
            qsCType.TopCount = 1;
            qsCType.AddOrder("hil_mdmtimestamp", OrderType.Descending);
            EntityCollection entCol = service.RetrieveMultiple(qsCType);
            if (entCol.Entities.Count > 0)
            {
                DateTime _cTimeStamp = entCol.Entities[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(330).AddSeconds(-30);
                if (_cTimeStamp.Year.ToString().PadLeft(4, '0') != "0001")
                    _enquiryDatetime = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + (_cTimeStamp.Second + 1).ToString().PadLeft(2, '0');
            }
            return _enquiryDatetime;
        }
        public static Guid CheckIfPartnerExists(string KUNNR, IOrganizationService service)
        {
            Guid Partner = new Guid();
            Partner = Guid.Empty;
            QueryExpression Query = new QueryExpression(Account.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, KUNNR);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                Partner = Found.Entities[0].Id;
            }
            return Partner;
        }
        #region GET INTEGRATION URL
        public static Integration GetIntegration(IOrganizationService service, string RecName)
        {
            try
            {
                Integration intConf = new Integration();
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, RecName);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                intConf.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                intConf.userName = integrationConfiguration.GetAttributeValue<string>("hil_username");
                intConf.passWord = integrationConfiguration.GetAttributeValue<string>("hil_password");
                return intConf;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
                throw new Exception("Error : " + ex.Message);
            }
        }
        #endregion
        #region App Setting Load/CRM Connection
        private static void LoadAppSettings()
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
                Console.WriteLine("ChannelPartnerSyncJob.Program.Main.LoadAppSettings ::  Error While Loading App Settings:" + ex.Message.ToString());
            }
        }
        private static void ConnectToCRM()
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
                Console.WriteLine("ChannelPartnerSyncJob.Program.Main.ConnectToCRM :: Error While Creating Connection with MS CRM Organisation:" + ex.Message.ToString());
            }
        }
        #endregion
    }
    public class Integration
    {
        public string uri { get; set; }
        public string userName { get; set; }
        public string passWord { get; set; }
    }
}
