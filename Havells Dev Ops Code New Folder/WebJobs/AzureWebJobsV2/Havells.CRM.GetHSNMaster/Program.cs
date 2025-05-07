using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.GetHSNMaster
{
    internal class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        static void Main(string[] args)
        {
            IOrganizationService _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                getHSNMaster(_service, (args.Length > 0 ? args[0] : ""));
            }
        }
        static void getHSNMaster(IOrganizationService service, string _syncDatetime)
        {
            try
            {
                Integration intConf = GetIntegration(service, "HSNCode");
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
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "Basic " + authInfo);
                Console.WriteLine("Downloading HSN Master Data: " + DateTime.Now.ToString());
                IRestResponse response = client.Execute(request);
                HSNMaster rootObject = JsonConvert.DeserializeObject<HSNMaster>(response.Content);
                Console.WriteLine("Downloading Completed of HSN Master Data: " + DateTime.Now.ToString());
                UpdateHSNMasterData(service, rootObject);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error " + ex.Message);
            }
        }
        public static void UpdateHSNMasterData(IOrganizationService _service, HSNMaster responseHSN)
        {
            int created = 0;
            int updated = 0;
            int deactivated = 0;
            int error = 0;
            try
            {
                if (responseHSN.Results.Count > 0)
                {
                    foreach (var item in responseHSN.Results)
                    {
                        try
                        {
                           // Console.WriteLine("HSN Code " + item.STEUC + " Started.");
                            Entity hsncode = new Entity("hil_hsncode");
                            hsncode["hil_name"] = item.STEUC;
                            hsncode["hil_taxtypetext"] = item.TEXT1;
                            hsncode["hil_taxpercentage"] = Convert.ToDecimal(item.KBETR);
                            hsncode["hil_mtimestamp"] = ConvertToDateTime(item.MTIMESTAMP);
                            //hsncode["hil_efffromdate"] = ConvertToDateTime(item.MTIMESTAMP);
                            //hsncode["hil_efftodate"] = ConvertToDateTime(item.MTIMESTAMP);

                            QueryExpression qrExp = new QueryExpression("hil_hsncode");
                            qrExp.ColumnSet = new ColumnSet(false);
                            qrExp.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, item.STEUC));
                            //qrExp.AddOrder("createdon", OrderType.Descending);
                            EntityCollection entCol = _service.RetrieveMultiple(qrExp);
                            if (entCol.Entities.Count > 0)
                            {
                                hsncode.Id = entCol[0].Id;
                                if (item.DELETE_FLAG.ToUpper() == "X")
                                {
                                    SetStateRequest req2 = new SetStateRequest();
                                    req2.State = new OptionSetValue(1);
                                    req2.Status = new OptionSetValue(2);
                                    req2.EntityMoniker = entCol[0].ToEntityReference();
                                    SetStateResponse res = (SetStateResponse)_service.Execute(req2);
                                    deactivated++;
                                    Console.WriteLine("HSN Code " + item.STEUC + " Deactivated.");
                                }
                                else
                                {
                                    SetStateRequest req2 = new SetStateRequest();
                                    req2.State = new OptionSetValue(0);
                                    req2.Status = new OptionSetValue(1);
                                    req2.EntityMoniker = entCol[0].ToEntityReference();
                                    SetStateResponse res = (SetStateResponse)_service.Execute(req2);
                                    _service.Update(hsncode);
                                    updated++;
                                    Console.WriteLine("HSN Code " + item.STEUC + " Updated.");
                                }
                            }
                            else
                            {
                                _service.Create(hsncode);
                                created++;
                                Console.WriteLine("HSN Code " + item.STEUC + " Created.");
                            }
                        }
                        catch (Exception ex)
                        {
                            error++;
                            Console.WriteLine("HSN Code " + item.STEUC + " has Error.  Error :- "+ ex.Message);
                        }
                    }
                    Console.WriteLine();
                    Console.WriteLine();
                    Console.WriteLine("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                    Console.WriteLine("Total Records " + responseHSN.Results.Count);
                    Console.WriteLine("Total Records Created " + created);
                    Console.WriteLine("Total Records Updated " + updated);
                    Console.WriteLine("Total Records Deactivated " + deactivated);
                    Console.WriteLine("Total Records Error " + error);
                    Console.WriteLine("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static string getTimeStamp(IOrganizationService service)
        {
            string _enquiryDatetime = "20210804000000";
            QueryExpression qsCType = new QueryExpression("hil_hsncode");
            qsCType.ColumnSet = new ColumnSet("hil_mtimestamp");
            qsCType.NoLock = true;
            qsCType.TopCount = 1;
            qsCType.AddOrder("hil_mtimestamp", OrderType.Descending);
            EntityCollection entCol = service.RetrieveMultiple(qsCType);
            if (entCol.Entities.Count > 0)
            {
                DateTime _cTimeStamp = entCol.Entities[0].GetAttributeValue<DateTime>("hil_mtimestamp").AddMinutes(330);
                if (_cTimeStamp.Year.ToString().PadLeft(4, '0') != "0001")
                    _enquiryDatetime = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + _cTimeStamp.Second.ToString().PadLeft(2, '0');
            }
            return _enquiryDatetime;
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
        #region App Setting Load/CRM Connection
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
        #endregion
    }
    public class HSNMaster
    {
        public string Result { get; set; }
        public List<HSNMasterModel> Results { get; set; }
    }
    public class HSNMasterModel
    {
        public string LAND1 { get; set; }
        public string STEUC { get; set; }
        public string TEXT1 { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CREATEDBY { get; set; }
        public string CTIMESTAMP { get; set; }
        public string MODIFYBY { get; set; }
        public string MTIMESTAMP { get; set; }
        public decimal KBETR { get; set; }

    }
    public class Integration
    {
        public string uri { get; set; }
        public string userName { get; set; }
        public string passWord { get; set; }
    }
}
