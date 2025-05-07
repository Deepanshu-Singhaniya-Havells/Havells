using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Text;

namespace Havells.CRM.ProductAndPriceListSync
{
    public class PriceListSync
    {
        static int iUpdateCount = 0;
        static int iCreateCount = 0;
        static int iDone = 0;
        static int iTotal = 0;
        public static void SparePartPriceList(IOrganizationService _service, string _syncDatetime, string motorGuid, string CableGUID)
        {
            Console.WriteLine("*******************************  PriceSync Started ******************************* ");
            if (_service != null)
            {
                Integration intConf = Models.GetIntegration(_service, "GetPriceDTLMaster");

                string _MTimeStamp = _syncDatetime;
                if (_syncDatetime == null || _syncDatetime == "")
                {
                    QueryExpression qsCType = new QueryExpression("hil_stagingpricingmapping");
                    qsCType.ColumnSet = new ColumnSet("hil_mdmtimestamp");
                    qsCType.NoLock = true;
                    qsCType.TopCount = 1;
                    qsCType.AddOrder("hil_mdmtimestamp", OrderType.Descending);
                    EntityCollection entCol = _service.RetrieveMultiple(qsCType);
                    if (entCol.Entities.Count > 0)
                    {
                        DateTime _cTimeStamp = entCol.Entities[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(300);
                        _MTimeStamp = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + _cTimeStamp.Second.ToString().PadLeft(2, '0');
                    }
                }
                else
                {
                    _MTimeStamp = _syncDatetime;
                }
                Console.WriteLine("TimeStam: " + _MTimeStamp);

                RequestPayload _payload = new RequestPayload() { FromDate = _MTimeStamp, IsInitialLoad = false };
                string data = JsonConvert.SerializeObject(_payload);
                Console.WriteLine("URL: " + intConf.uri);
                Console.WriteLine("Request: " + data);


                string _authInfo = intConf.userName + ":" + intConf.passWord;

                Console.WriteLine("authInfo: " + _authInfo);
                _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));

                var client = new RestClient(intConf.uri);
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
                    iTotal = rootObject.Results.Count;
                    Console.WriteLine("iTotal Records: " + iTotal);
                    if (iTotal > 0)
                    {
                        foreach (PricingResult obj in rootObject.Results)
                        {
                            // if (obj.MATNR == "CHCAJLDFFK202X5")//|| obj.MATNR == "CODQMLDWZK041X5"|| obj.MATNR == "COCQMLDWZK021X5" || obj.MATNR == "COCQMLDWZK041X5" || obj.MATNR == "COCQMLDWZK061X5")
                            {
                                try
                                {
                                    if (obj.KSCHL == "ZWEB")
                                    {
                                        Guid productDivision = Models.GetProductDivision(obj.MATNR, _service);
                                        if (productDivision != new Guid(motorGuid) && productDivision != Guid.Empty && productDivision != new Guid(CableGUID))//not equal to Motor and Cable
                                        {
                                            CreateUpdatePriceListStaging(_service, obj);
                                            iDone++;
                                        }
                                    }
                                    else if (obj.KSCHL == "ZPR0")
                                    {
                                        Guid productDivision = Models.GetProductDivision(obj.MATNR, _service);
                                        if (productDivision == new Guid(motorGuid) || productDivision == new Guid(CableGUID))// equal to Motor or Cable
                                        {
                                            CreateUpdatePriceListStaging(_service, obj);
                                            iDone++;
                                        }
                                    }
                                    else
                                        Console.WriteLine("Price List Skiped For " + obj.MATNR + " Price is " + obj.KSCHL);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                    continue;
                                }
                            }
                        }
                        Console.WriteLine("Price List Updated with "+ iDone+" record ");
                    }
                    else
                    {
                        Console.WriteLine("Price List Updated with 0 record ");
                    }
                }
            }
            Console.WriteLine("*******************************  PriceSync Ended ******************************* ");
        }
        static void CreateUpdatePriceListStaging(IOrganizationService _service, PricingResult obj)
        {
            iDone += 1;
            Guid stagingProductPriceId = Models.GetStagingProductPrice(obj.MATNR, Convert.ToDateTime(obj.DATAB), Convert.ToDateTime(obj.DATBI), _service);
            if (stagingProductPriceId == Guid.Empty)
            {

                Guid ProductId = Models.GetProduct(obj.MATNR, _service);
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
                    iPartDivMapp["hil_mdmtimestamp"] = Models.StringToDateTime(obj.CTIMESTAMP);
                }
                else
                {
                    iPartDivMapp["hil_mdmtimestamp"] = Models.StringToDateTime(obj.MTIMESTAMP);
                }
                iPartDivMapp["hil_deleteflag"] = obj.DELETE_FLAG.ToString();
                _service.Create(iPartDivMapp);
                iCreateCount += 1;
                Console.WriteLine("Price List Created For " + obj.MATNR + " : with Amount " + obj.KBETR + " : " + iCreateCount.ToString() + "/" + iTotal.ToString());
            }
            else
            {

                Guid ProductId = Models.GetProduct(obj.MATNR, _service);
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
                    iPartDivMapp["hil_mdmtimestamp"] = Models.StringToDateTime(obj.CTIMESTAMP);
                }
                else
                {
                    iPartDivMapp["hil_mdmtimestamp"] = Models.StringToDateTime(obj.MTIMESTAMP);
                }
                _service.Update(iPartDivMapp);
                iUpdateCount += 1;
                Console.WriteLine("Price List Updated For " + obj.MATNR + " : with Amount " + obj.KBETR + " : " + iUpdateCount.ToString() + "/" + iTotal.ToString());

            }
        }
        public static void AMCPriceList(IOrganizationService _service)
        {
            Console.WriteLine("******************************* AMCPriceList Started ******************************* ");
            if (_service != null)
            {
                try
                {
                    Integration intConf = Models.GetIntegration(_service, "D365_AMC_DISC_PRC_DTL");

                    //var client = new RestClient("https://middleware.havells.com:50001/RESTAdapter/dynamics/D365_AMC_DISC_PRC_DTL");
                    var client = new RestClient(intConf.uri);// "https://p90ci.havells.com:50001/RESTAdapter/dynamics/D365_AMC_DISC_PRC_DTL");
                    client.Timeout = -1;

                    string _authInfo = intConf.userName + ":" + intConf.passWord;
                    //string _authInfo = "D365_Havells" + ":" + "PRDD365@1234";
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
                            Guid ProductId = Models.CheckForAMCProduct(row.MATNR, _service);
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
                                        _stageDev["hil_mdmtimestamp"] = Models.StringToDateTime(row.CTIMESTAMP);
                                    }
                                    else
                                    {
                                        _stageDev["hil_mdmtimestamp"] = Models.StringToDateTime(row.MTIMESTAMP);
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
                            //counter--;
                        }

                    }
                    client = new RestClient(intConf.uri);// "https://p90ci.havells.com:50001/RESTAdapter/dynamics/D365_AMC_DISC_PRC_DTL");
                    client.Timeout = -1;

                    _authInfo = intConf.userName + ":" + intConf.passWord; //_authInfo = "D365_Havells" + ":" + "PRDD365@1234";
                    _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));

                    request = new RestRequest(Method.POST);
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
            Console.WriteLine("*******************************  AMCPriceList Ended *******************************");
        }
    }
}
