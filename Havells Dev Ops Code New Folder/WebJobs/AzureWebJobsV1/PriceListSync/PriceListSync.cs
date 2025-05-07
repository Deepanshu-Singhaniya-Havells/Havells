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
using System.Collections.Generic;
using RestSharp;
using Newtonsoft.Json;

namespace PriceListSync
{
    public class PriceListSync
    {
        public static void SparePartPriceList(IOrganizationService _service) {
            if (_service != null) {
                string _MTimeStamp = "19000101000000";
                QueryExpression qsCType = new QueryExpression("hil_stagingpricingmapping");
                qsCType.ColumnSet = new ColumnSet("hil_mdmtimestamp");
                qsCType.NoLock = true;
                qsCType.TopCount = 1;
                qsCType.AddOrder("hil_mdmtimestamp", OrderType.Descending);
                EntityCollection entCol = _service.RetrieveMultiple(qsCType);
                if (entCol.Entities.Count > 0) 
                {
                    DateTime _cTimeStamp = entCol.Entities[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(330).AddSeconds(1);
                    _MTimeStamp = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + _cTimeStamp.Second.ToString().PadLeft(2, '0');
                }
                
                RequestPayload _payload = new RequestPayload() { FromDate = _MTimeStamp, IsInitialLoad = false, Condition = "ZWEB" };
                string data = JsonConvert.SerializeObject(_payload);

                string _authInfo = "D365_Havells" + ":" + "PRDD365@1234";
                _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));

                //var client = new RestClient("https://middleware.havells.com:50001/RESTAdapter/MDMService/Core/Product/GetPriceDTLMaster");
                var client = new RestClient("https://p90ci.havells.com:50001/RESTAdapter/MDMService/Core/Product/GetPriceDTLMaster");
                var request = new RestRequest(Method.POST);
                request.AddHeader("authorization", "Basic " + _authInfo);
                request.AddHeader("content-type", "application/json");
                request.AddParameter("application/json", data, ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                if (response.StatusCode != System.Net.HttpStatusCode.OK)
                {
                    Console.WriteLine(response.Content);
                }
                else
                {
                    var rootObject = Newtonsoft.Json.JsonConvert.DeserializeObject<PricingRoot>(response.Content);
                    int iDone = 0;
                    int iTotal = rootObject.Results.Count;
                    int iUpdateCount = 0;
                    int iCreateCount = 0;

                    if (iTotal > 0)
                    {
                        foreach (PricingResult obj in rootObject.Results)
                        {
                            try
                            {
                                //Price List Updated For CHLALLNAFK01300 : 10279/24599
                                iDone += 1;
                                Guid stagingProductPriceId = GetStagingProductPrice(obj.MATNR, Convert.ToDateTime(obj.DATAB), Convert.ToDateTime(obj.DATBI), _service);
                                if (obj.MATNR.IndexOf("CHSECTNNNK01070") >= 0)
                                {
                                    Console.WriteLine(obj.MATNR + "/" + obj.MTIMESTAMP);
                                }
                                if (stagingProductPriceId == Guid.Empty)
                                {
                                    Guid ProductId = GetProduct(obj.MATNR, _service);
                                    Entity iPartDivMapp = new Entity("hil_stagingpricingmapping");
                                    iPartDivMapp["hil_name"] = obj.MATNR;
                                    if (ProductId != Guid.Empty)
                                    {
                                        EntityReference iProdCat = new EntityReference("product", ProductId);
                                        iPartDivMapp["hil_product"] = (EntityReference)iProdCat;
                                    }
                                    iPartDivMapp["hil_price"] = Convert.ToInt32(obj.KBETR);
                                    iPartDivMapp["hil_datestart"] = Convert.ToDateTime(obj.DATAB);
                                    iPartDivMapp["hil_dateend"] = Convert.ToDateTime(obj.DATBI);
                                    if (obj.MTIMESTAMP == null)
                                    {
                                        iPartDivMapp["hil_mdmtimestamp"] = StringToDateTime(obj.CTIMESTAMP);
                                    }
                                    else
                                    {
                                        iPartDivMapp["hil_mdmtimestamp"] = StringToDateTime(obj.MTIMESTAMP);
                                    }
                                    iPartDivMapp["hil_deleteflag"] = obj.DELETE_FLAG.ToString();
                                    _service.Create(iPartDivMapp);
                                    iCreateCount += 1;
                                    Console.WriteLine("Price List Created For " + obj.MATNR + " : " + iCreateCount.ToString() + "/" + iTotal.ToString());
                                }
                                else
                                {
                                    Guid ProductId = GetProduct(obj.MATNR, _service);
                                    Entity iPartDivMapp = new Entity("hil_stagingpricingmapping");
                                    if (ProductId != Guid.Empty)
                                    {
                                        EntityReference iProdCat = new EntityReference("product", ProductId);
                                        iPartDivMapp["hil_product"] = (EntityReference)iProdCat;
                                    }
                                    iPartDivMapp["hil_price"] = Convert.ToInt32(obj.KBETR);
                                    iPartDivMapp.Id = stagingProductPriceId;
                                    iPartDivMapp["hil_deleteflag"] = obj.DELETE_FLAG.ToString();
                                    if (obj.MTIMESTAMP == null)
                                    {
                                        iPartDivMapp["hil_mdmtimestamp"] = StringToDateTime(obj.CTIMESTAMP);
                                    }
                                    else
                                    {
                                        iPartDivMapp["hil_mdmtimestamp"] = StringToDateTime(obj.MTIMESTAMP);
                                    }
                                    _service.Update(iPartDivMapp);
                                    iUpdateCount += 1;
                                    Console.WriteLine("Price List Updated For " + obj.MATNR + " : " + iUpdateCount.ToString() + "/" + iTotal.ToString());
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                                continue;
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("Price List Updated with 0 record ");
                    }
                }
            }
        }
        public static void AMCPriceList(IOrganizationService _service)
        {
            if (_service != null)
            {
                try
                {
                    //var client = new RestClient("https://middleware.havells.com:50001/RESTAdapter/dynamics/D365_AMC_DISC_PRC_DTL");
                    var client = new RestClient("https://p90ci.havells.com:50001/RESTAdapter/dynamics/D365_AMC_DISC_PRC_DTL");
                    client.Timeout = -1;

                    string _authInfo = "D365_Havells" + ":" + "PRDD365@1234";
                    _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));

                    var request = new RestRequest(Method.POST);
                    request.AddHeader("authorization", "Basic " + _authInfo);
                    request.AddHeader("Content-Type", "application/json");
                    request.AddParameter("application/json", "{\r\n    \"IM_FLAG\": \"R\"\r\n}", ParameterType.RequestBody);
                    IRestResponse response = client.Execute(request);
                    //Console.WriteLine(response.Content);
                    List<LTTABLE> table = new List<LTTABLE>();
                    table = JsonConvert.DeserializeObject<OutputClass>(response.Content).LT_TABLE;
                    Console.WriteLine("Totla Count: " + table.Count);
                    var i = 1;
                    int counter = 0;
                    for (; table.Count > counter;)
                    {
                        LTTABLE row = table[counter];
                        //    foreach (LTTABLE row in table)
                        //{
                        try
                        {
                            Guid ProductId = CheckForAMCProduct(row.MATNR, _service);
                            if (ProductId != Guid.Empty)
                            {

                                QueryExpression Query = new QueryExpression("hil_stagingpricingmapping");
                                Query.ColumnSet = new ColumnSet(true);
                                Query.Criteria = new FilterExpression(LogicalOperator.And);
                                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, row.MATNR);
                                Query.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, row.VKORG);
                                Query.Criteria.AddCondition("hil_price", ConditionOperator.Equal, (int)(float.Parse(row.KBETR)));
                                Query.Criteria.AddCondition("hil_type", ConditionOperator.Equal, row.KSCHL == "ZPR0" ? true : false);
                                Query.Criteria.AddCondition("hil_datestart", ConditionOperator.Equal, Convert.ToDateTime(row.DATBI));
                                Query.Criteria.AddCondition("hil_dateend", ConditionOperator.Equal, Convert.ToDateTime(row.DATAB));
                                EntityCollection Found = _service.RetrieveMultiple(Query);
                                if (Found.Entities.Count == 0)
                                {
                                    Entity _stageDev = new Entity("hil_stagingpricingmapping");
                                    _stageDev["hil_salesoffice"] = row.VKORG;
                                    _stageDev["hil_name"] = row.MATNR;
                                    _stageDev["hil_type"] = row.KSCHL == "ZPR0" ? true : false;
                                    _stageDev["hil_price"] = (int)(float.Parse(row.KBETR));
                                    _stageDev["hil_datestart"] = Convert.ToDateTime(row.DATBI);
                                    _stageDev["hil_dateend"] = Convert.ToDateTime(row.DATAB);
                                    _stageDev["hil_deleteflag"] = row.DELETE_FLAG.ToString();

                                    if (row.MTIMESTAMP == null)
                                    {
                                        _stageDev["hil_mdmtimestamp"] = StringToDateTime(row.CTIMESTAMP);
                                    }
                                    else
                                    {
                                        _stageDev["hil_mdmtimestamp"] = StringToDateTime(row.MTIMESTAMP);
                                    }
                                    _stageDev["hil_product"] = new EntityReference("product", ProductId);

                                    _service.Create(_stageDev);

                                    Console.WriteLine(i + " Record Created " + row.MATNR);
                                }
                                decimal _amount = 0;
                                decimal _discPer = 0;
                                if (row.KSCHL == "ZPR0") //Price
                                {
                                    _amount = decimal.Parse(row.KBETR);
                                }
                                else
                                {  //Discount
                                    _discPer = decimal.Parse(row.KBETR);
                                }
                                Query = new QueryExpression("hil_stagingpricingmapping");
                                Query.ColumnSet = new ColumnSet("hil_type", "hil_price");
                                Query.Criteria = new FilterExpression(LogicalOperator.And);
                                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, row.MATNR);
                                Query.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, row.VKORG);
                                Query.Criteria.AddCondition("hil_type", ConditionOperator.Equal, row.KSCHL == "ZPR0" ? false : true);
                                Query.Criteria.AddCondition("hil_datestart", ConditionOperator.OnOrAfter, DateTime.Now);
                                Query.Criteria.AddCondition("hil_dateend", ConditionOperator.OnOrBefore, DateTime.Now);
                                EntityCollection entColl = _service.RetrieveMultiple(Query);
                                if (entColl.Entities.Count > 0)
                                {
                                    if (entColl.Entities[0].GetAttributeValue<bool>("hil_type")) //Price
                                    {
                                        _amount = decimal.Parse(entColl.Entities[0].GetAttributeValue<int>("hil_price").ToString());
                                    }
                                    else
                                    {  //Discount
                                        _discPer = decimal.Parse(entColl.Entities[0].GetAttributeValue<int>("hil_price").ToString());
                                    }
                                }
                                if (_amount > 0)
                                {
                                    if (_discPer > 0)
                                    {
                                        _amount = Math.Round(_amount - (_amount * _discPer / 100), 0);
                                    }
                                    Entity entProd = new Entity("product", ProductId);
                                    entProd["hil_amount"] = new Money(_amount);
                                    _service.Update(entProd);
                                }
                            }
                            else
                            {
                                Console.WriteLine("Non AMC Product: " + i.ToString() + "/" + table.Count);
                            }
                            i++;
                            counter++;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error " + ex.Message);
                            try
                            {
                                Program.loginUserGuid = ((WhoAmIResponse)_service.Execute(new WhoAmIRequest())).UserId;
                            }
                            catch
                            {
                                try
                                {
                                    ClientCredentials credentials = new ClientCredentials();
                                    credentials.UserName.UserName = Program._userId;
                                    credentials.UserName.Password = Program._password;
                                    Uri serviceUri = new Uri(Program._soapOrganizationServiceUri);
                                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                                    ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                                    OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);
                                    proxy.EnableProxyTypes();
                                    _service = (IOrganizationService)proxy;
                                    Program.loginUserGuid = ((WhoAmIResponse)_service.Execute(new WhoAmIRequest())).UserId;
                                }
                                catch
                                {
                                    Console.WriteLine("Service not created....");
                                }
                                counter--;
                            }
                            //counter--;
                        }
                        
                    }
                    client = new RestClient("https://p90ci.havells.com:50001/RESTAdapter/dynamics/D365_AMC_DISC_PRC_DTL");
                    client.Timeout = -1;

                    _authInfo = "D365_Havells" + ":" + "PRDD365@1234";
                    _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));

                    request = new RestRequest(Method.GET);
                    request.AddHeader("authorization", "Basic " + _authInfo);
                    request.AddHeader("Content-Type", "application/json");
                    request.AddParameter("application/json", "{\r\n \"IM_FLAG\": \"U\"\r\n}", ParameterType.RequestBody);
                    response = client.Execute(request);
                    Console.WriteLine("Done");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error  " + ex.Message);
                    Console.ReadLine();
                }

            }
        }
        public static DateTime? StringToDateTime(string _mdmTimeStamp ) {
            DateTime? _dtMDMTimeStamp = null;
            try
            {
                _dtMDMTimeStamp = new DateTime(Convert.ToInt32(_mdmTimeStamp.Substring(0,4)), Convert.ToInt32(_mdmTimeStamp.Substring(4, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(6, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(8, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(10, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(12, 2)));
            }
            catch { }
            return _dtMDMTimeStamp;
        }
        public static Guid GetProduct(string materialcode, IOrganizationService service)
        {
            Guid iDivision = new Guid();
            iDivision = Guid.Empty;
            QueryExpression Query = new QueryExpression("product");
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("name", ConditionOperator.Equal, materialcode);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                iDivision = Found.Entities[0].Id;
            }
            return iDivision;
        }

        public static Guid CheckForAMCProduct(string materialcode, IOrganizationService service)
        {
            Guid iProduct = Guid.Empty;
            QueryExpression Query = new QueryExpression("product");
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("name", ConditionOperator.Equal, materialcode);
            Query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 910590001);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                iProduct = Found.Entities[0].Id;
            }
            return iProduct;
        }

        public static Guid GetStagingProductPrice(string materialcode, DateTime startDate, DateTime endDate, IOrganizationService service)
        {
            Guid iDivision = new Guid();
            iDivision = Guid.Empty;
            QueryExpression Query = new QueryExpression("hil_stagingpricingmapping");
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, materialcode);
            Query.Criteria.AddCondition("hil_datestart", ConditionOperator.Equal, startDate);
            Query.Criteria.AddCondition("hil_dateend", ConditionOperator.Equal, endDate);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                iDivision = Found.Entities[0].Id;
            }
            return iDivision;
        }
    }

    public class LTTABLE
    {
        public string VKORG { get; set; }
        public string MATNR { get; set; }
        public string VKBUR { get; set; }
        public string KSCHL { get; set; }
        public string KBETR { get; set; }
        public string KONWA { get; set; }
        public string DATBI { get; set; }
        public string DATAB { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CREATEDBY { get; set; }
        public string CTIMESTAMP { get; set; }
        public string MODIFYBY { get; set; }
        public string MTIMESTAMP { get; set; }
    }
    public class OutputClass
    {
        public string EV_RETURN { get; set; }
        public List<LTTABLE> LT_TABLE { get; set; }
    }


    public class RequestPayload
    {
        public string FromDate { get; set; }
        public string ToDate { get; set; }
        public bool IsInitialLoad { get; set; }
        public string Condition { get; set; }

    }
    public class PricingRoot
    {
        public object Result { get; set; }
        public List<PricingResult> Results { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }

    public class PricingResult
    {
        public string MATNR { get; set; }
        public decimal KBETR { get; set; }
        public string DATAB { get; set; }
        public string DATBI { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CREATEDBY { get; set; }
        public string MODIFYBY { get; set; }
        public string CTIMESTAMP { get; set; }
        public string MTIMESTAMP { get; set; }
    }
}
