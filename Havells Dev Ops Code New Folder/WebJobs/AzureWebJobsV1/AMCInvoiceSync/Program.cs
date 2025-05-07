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
using System.Collections.Generic;
using System.Globalization;
using RestSharp;
using Microsoft.Xrm.Tooling.Connector;

namespace AMCInvoiceSync
{
    class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                SyncSAPAMCInvoiceData();
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
        static void SyncSAPAMCInvoiceData()
        {
            try
            {
                if (_service != null)
                {
                    String DynamicsURL = "https://p90ci.havells.com:50001/RESTAdapter/dynamics/D365_AMC_BILLING_DETAIL";
                    String fromDate = GetLastAMCPurchaseDate();
                    String endDate = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0');
                    string _authInfo = "D365_Havells" + ":" + "PRDD365@1234";
                    _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));

                    RequestData data = new RequestData();
                    data.FROM_DATE = fromDate;
                    data.TO_DATE = endDate;
                    data.IM_FLAG = "R";
                    String requestStr = JsonConvert.SerializeObject(data);
                    var client = new RestClient(DynamicsURL);
                    client.Timeout = -1;
                    var request = new RestRequest(Method.POST);
                    request.AddHeader("Authorization", "Basic " + _authInfo);
                    request.AddHeader("Content-Type", "application/json");

                    request.AddParameter("application/json", requestStr, ParameterType.RequestBody);
                    IRestResponse response = client.Execute(request);

                    var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<ResponseData>(response.Content);
                    if (obj.RETURN != "No data found")
                    {
                        int rowCount = 1;
                        int totalCount = obj.LT_TABLE.Count;
                        foreach (table table in obj.LT_TABLE)
                        {
                            try
                            {
                                CultureInfo culture = new CultureInfo("en-US");
                                EntityReference entAMCPlan = null;
                                DateTime amcStartDate = Convert.ToDateTime(table.AMCSTART, culture);
                                DateTime amcEndDate = Convert.ToDateTime(table.AMCEND, culture);
                                QueryExpression query = new QueryExpression("product");
                                query.ColumnSet = new ColumnSet(false);
                                query.Criteria = new FilterExpression(LogicalOperator.And);
                                query.Criteria.AddCondition("productnumber", ConditionOperator.Equal, table.MATNR);
                                EntityCollection entCol = _service.RetrieveMultiple(query);
                                if (entCol.Entities.Count > 0)
                                {
                                    entAMCPlan = entCol[0].ToEntityReference();
                                    if (GetStagingAMCInvoiceData(entAMCPlan, table.SERIAL, amcStartDate, amcEndDate, _service) == Guid.Empty && entAMCPlan != null)
                                    {
                                        Entity entity = new Entity("hil_amcstaging");
                                        entity["hil_serailnumber"] = table.SERIAL;
                                        entity["hil_amcplan"] = entAMCPlan;
                                        entity["hil_name"] = table.VBELN_B.ToString();
                                        DateTime amcPurchaseDate = Convert.ToDateTime(table.CUSTPDATE, culture);
                                        DateTime amcBillingDate = Convert.ToDateTime(table.ERDAT, culture);
                                        entity["hil_sapbillingdate"] = amcBillingDate;
                                        entity["hil_warrantystartdate"] = amcStartDate;//date
                                        entity["hil_warrantyenddate"] = amcEndDate;//date
                                        entity["hil_amcpurchasedate"] = amcPurchaseDate;//date
                                        _service.Create(entity);
                                        Console.WriteLine(rowCount.ToString() + "/" + totalCount.ToString() + " Serial Number: " + table.SERIAL);
                                    }
                                    else {
                                        Console.WriteLine(rowCount.ToString() + "/" + totalCount.ToString() + "AMC Invoice already exists.");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(rowCount.ToString() + "/" + totalCount.ToString() + " Error!!! " + ex.Message);
                            }
                            rowCount += 1;
                        }
                        Console.WriteLine("DONE");
                    }
                }
                else {
                    Console.WriteLine("Error while establishing connection with D365.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error  " + ex.Message);
            }
        }
        static string GetLastAMCPurchaseDate()
        {
            string _lastAMCPurchaseDate = "2021-01-01";
            try
            {
                QueryExpression Query = new QueryExpression("hil_amcstaging");
                Query.ColumnSet = new ColumnSet("hil_amcpurchasedate");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_amcpurchasedate", ConditionOperator.NotNull);
                Query.TopCount = 1;
                Query.AddOrder("hil_amcpurchasedate", OrderType.Descending);

                EntityCollection Found = _service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    DateTime _date = Found.Entities[0].GetAttributeValue<DateTime>("hil_amcpurchasedate").AddMinutes(330);
                    _lastAMCPurchaseDate = _date.Year.ToString() + "-" + _date.Month.ToString().PadLeft(2, '0') + "-" + _date.Day.ToString().PadLeft(2,'0');
                }
                return _lastAMCPurchaseDate;
            }
            catch {}
            return _lastAMCPurchaseDate;
        }
        static Guid GetStagingAMCInvoiceData(EntityReference amcPlan, string serialNumber, DateTime startDate, DateTime endDate, IOrganizationService service)
        {
            Guid _retGuid = Guid.Empty;
            QueryExpression Query = new QueryExpression("hil_amcstaging");
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_amcplan", ConditionOperator.Equal, amcPlan.Id);
            Query.Criteria.AddCondition("hil_serailnumber", ConditionOperator.Equal, serialNumber);
            Query.Criteria.AddCondition("hil_warrantystartdate", ConditionOperator.Equal, startDate);
            Query.Criteria.AddCondition("hil_warrantyenddate", ConditionOperator.Equal, endDate);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                _retGuid = Found.Entities[0].Id;
            }
            return _retGuid;
        }
    }
    public class RequestData
    {
        public string IM_FLAG { get; set; }
        public string FROM_DATE { get; set; }
        public string TO_DATE { get; set; }
    }
    public class table
    {
        public string SERIAL { get; set; }
        public string COUNTER { get; set; }
        public string CALLNO { get; set; }
        public string MOBILENO { get; set; }
        public string KUNNR { get; set; }
        public string VBELN_S { get; set; }
        public string POSNR_S { get; set; }
        public long VBELN_B { get; set; }
        public string POSNR_B { get; set; }
        public string MATNR { get; set; }
        public string AMCSTART { get; set; }
        public string AMCEND { get; set; }
        public string CUSTPDATE { get; set; }
        public string BSTDK { get; set; }
        public string DEL_FLAG { get; set; }
        public string ERDAT { get; set; }
    }
    public class ResponseData
    {
        public string RETURN { get; set; }
        public List<table> LT_TABLE { get; set; }
    }

}
