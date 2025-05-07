using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

namespace Havells.Crm.PriceListSync
{
    public class Program
    {
        public static int iDone = 0;
        public static int skip = 0;
        public static int iUpdateCount = 0;
        public static int iDeleted = 0;
        public static int iCreateCount = 0;
        public static int iTotal = 0;
        public static string mtimeStam = "";
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = HavellsConnection.CreateConnection.createConnection(finalString);

           
            priceListItemSync(service);
        }
        public static void updatePriceInStagging(IOrganizationService _service, PricingResult obj)
        {
            try
            {
                //Price List Updated For CHLALLNAFK01300 : 10279/24599
                iDone += 1;
                Guid stagingProductPriceId = GetStagingProductPrice(obj.MATNR, Convert.ToDateTime(obj.DATAB), Convert.ToDateTime(obj.DATBI), _service);
                //if (obj.MATNR.IndexOf("CHSECTNNNK01070") >= 0)
                //{
                //    Console.WriteLine(obj.MATNR + "/" + obj.MTIMESTAMP);
                //}
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
            }
        }
        public static DateTime? StringToDateTime(string _mdmTimeStamp)
        {
            DateTime? _dtMDMTimeStamp = null;
            try
            {
                _dtMDMTimeStamp = new DateTime(Convert.ToInt32(_mdmTimeStamp.Substring(0, 4)), Convert.ToInt32(_mdmTimeStamp.Substring(4, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(6, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(8, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(10, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(12, 2)));
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
       //startup function
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
        public static void priceListItemSync(IOrganizationService _service)
        {
            if (_service != null)
            {
                IntegrationConfig intConfig = Model.IntegrationConfiguration(_service, "GetPriceDTLMaster");
                string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));

                string _MTimeStamp = intConfig.MTIMESTAMP;
                mtimeStam = _MTimeStamp;
                RequestPayload _payload = new RequestPayload()
                {
                    FromDate = _MTimeStamp,
                    IsInitialLoad = false,
                    Condition = "ZWEB"
                };
                string data = JsonConvert.SerializeObject(_payload);

                var client = new RestClient(intConfig.uri);
                var request = new RestRequest(Method.POST);
                request.AddHeader("authorization", authInfo);
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
                    int iTotal = rootObject.Results.Count;
                    if (iTotal > 0)
                    {
                        Console.WriteLine("************************************ Price List Sync Started ********************************\n");
                        Console.WriteLine("\t\t******************************Total Record Count =\t" + iTotal.ToString() + "\t********************************");
                        APIResponse(rootObject, _service);
                    }
                    else
                    {
                        Console.WriteLine("Price List Updated with 0 record ");
                    }
                }
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet(false);
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "GetPriceDTLMaster");
                Entity integrationConfigurationUpdate = _service.RetrieveMultiple(qsCType)[0];
                integrationConfigurationUpdate["hil_contactno"] = mtimeStam;
                _service.Update(integrationConfigurationUpdate);
            }
        }
        public static void APIResponse(PricingRoot rootObject, IOrganizationService _service)
        {
            iTotal = rootObject.Results.Count;

            foreach (PricingResult obj in rootObject.Results)
            {
                if (obj.MTIMESTAMP == null)
                {
                    mtimeStam = obj.CTIMESTAMP;
                }
                else
                {
                    mtimeStam = obj.MTIMESTAMP;
                }

                try
                {
                    iDone += 1;
                    List<Guid> ProductDivision = GetProductAndDivision(obj.MATNR, _service);
                    if (ProductDivision.Count > 0)
                    {
                        Guid DivisionId = ProductDivision[1];
                        Guid ProductId = ProductDivision[0];
                        if (ProductId != Guid.Empty && DivisionId != Guid.Empty)
                        {
                            Guid iPriceList = Guid.Empty;
                            Guid Uom = Guid.Empty;
                            int convirsionFactor = 1;
                            QueryExpression Query = new QueryExpression("hil_enquirydepartment");
                            Query.ColumnSet = new ColumnSet("hil_pricelist", "hil_uom", "hil_conversionfactor");
                            Query.Criteria.AddCondition("hil_productdivision", ConditionOperator.Equal, DivisionId);
                            EntityCollection Found = _service.RetrieveMultiple(Query);
                            if (Found.Entities.Count == 1)
                            {
                                iPriceList = Found[0].GetAttributeValue<EntityReference>("hil_pricelist").Id;
                                Uom = Found[0].GetAttributeValue<EntityReference>("hil_uom").Id;
                                convirsionFactor = Found[0].GetAttributeValue<int>("hil_conversionfactor");
                                decimal lp = obj.KBETR / convirsionFactor;
                                QueryExpression query = new QueryExpression("productpricelevel");
                                query.ColumnSet = new ColumnSet("amount", "uomid", "productpricelevelid");
                                query.Distinct = true;
                                query.Criteria.AddCondition("productid", ConditionOperator.Equal, ProductId);
                                query.Criteria.AddCondition("pricelevelid", ConditionOperator.Equal, iPriceList);
                                EntityCollection pricelistItemColl = _service.RetrieveMultiple(query);
                                if (pricelistItemColl.Entities.Count == 1)
                                {
                                    if (obj.DELETE_FLAG.ToLower() != "X".ToLower())
                                    {
                                        Entity pricelistItem = new Entity("productpricelevel");
                                        pricelistItem.Id = pricelistItemColl[0].Id;
                                        pricelistItem["amount"] = new Money(lp);
                                        _service.Update(pricelistItem);
                                        iUpdateCount += 1;
                                        Console.WriteLine("Price List Updated For PriceListItem " + obj.MATNR + " : " + iUpdateCount.ToString() + "/" + iTotal.ToString());

                                    }
                                    else
                                    {
                                        iDeleted++;
                                        _service.Delete(pricelistItemColl.EntityName, pricelistItemColl[0].Id);
                                        Console.WriteLine("Price List Deleted For PriceListItem " + obj.MATNR + " : " + iDeleted.ToString() + "/" + iTotal.ToString());
                                    }
                                }
                                else
                                {
                                    if (obj.DELETE_FLAG.ToLower() != "X".ToLower())
                                    {
                                        Entity pricelistItem = new Entity("productpricelevel");
                                        pricelistItem["productid"] = new EntityReference("product", ProductId);
                                        pricelistItem["amount"] = new Money(lp);
                                        pricelistItem["pricelevelid"] = new EntityReference("pricelevel", iPriceList);
                                        pricelistItem["uomid"] = new EntityReference("uom", Uom);
                                        pricelistItem["quantitysellingcode"] = new OptionSetValue(1);
                                        _service.Create(pricelistItem);
                                        iCreateCount += 1;
                                        Console.WriteLine("Price List Created For PriceListItem " + obj.MATNR + " : " + iCreateCount.ToString() + "/" + iTotal.ToString());
                                    }
                                }
                            }
                            else
                            {
                                // Department not fond then existing Code need to implemented.
                                updatePriceInStagging(_service, obj);
                            }
                        }
                        else
                        {
                            //division not found
                            skip++;
                            Console.WriteLine("Price List Skip For PriceListItem " + obj.MATNR + " : " + skip.ToString() + "/" + iTotal.ToString());
                        }
                    }
                    else
                    {
                        //product is not found
                        skip++;
                        Console.WriteLine("Price List Skip For PriceListItem " + obj.MATNR + " : " + skip.ToString() + "/" + iTotal.ToString());
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error  For PriceListItem " + obj.MATNR + " : " + ex.Message);
                    continue;
                }
            }
        }
        public static List<Guid> GetProductAndDivision(string materialcode, IOrganizationService service)
        {
            List<Guid> iDivision = new List<Guid>();

            QueryExpression Query = new QueryExpression("product");
            Query.ColumnSet = new ColumnSet("hil_division");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("name", ConditionOperator.Equal, materialcode);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                iDivision.Add(Found.Entities[0].Id);
                iDivision.Add(Found[0].Contains("hil_division") ? Found[0].GetAttributeValue<EntityReference>("hil_division").Id : Guid.Empty);
            }
            return iDivision;
        }
    }
}
