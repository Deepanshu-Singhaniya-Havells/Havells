using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace Havells.CRM.PriceListItemCreation
{
    class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        public static int done = 0;
        public static int totlal = 0;
        static void Main(string[] args)
        {
            IOrganizationService service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((CrmServiceClient)service).IsReady)
            {
                syncPriceListItem(service);
            }
        }
        private static IOrganizationService ConnectToCRM(string connectionString)
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
        private static void syncPriceListItem(IOrganizationService service)
        {
            int count = 0;
            QueryExpression qProduct = new QueryExpression("product");
            qProduct.ColumnSet = new ColumnSet("hil_division", "hil_amount");
            qProduct.Criteria = new FilterExpression(LogicalOperator.And);
            qProduct.Criteria.AddCondition("hil_amount", ConditionOperator.GreaterThan, 0);
            qProduct.Criteria.AddCondition("hil_division", ConditionOperator.Equal, new Guid("FD555381-16FA-E811-A94D-000D3AF06CD4"));//HAVELLS CABLE
            qProduct.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 5);//Material
            qProduct.Criteria.AddCondition("producttypecode", ConditionOperator.Equal, 1);//Finished Goods
            qProduct.AddOrder("createdon", OrderType.Ascending);
            qProduct.PageInfo = new PagingInfo();
            qProduct.PageInfo.Count = 5000;
            qProduct.PageInfo.PageNumber = 1;
            qProduct.PageInfo.ReturnTotalRecordCount = true;
            try
            {
                EntityCollection productColl = service.RetrieveMultiple(qProduct);
                do
                {
                    if (productColl.Entities.Count > 0)
                    {
                        totlal = totlal + productColl.Entities.Count;
                        foreach (Entity ent in productColl.Entities)
                        {
                            createUpdatePriceListItem(ent, service);
                        }
                    }
                    qProduct.PageInfo.PageNumber += 1;
                    qProduct.PageInfo.PagingCookie = productColl.PagingCookie;
                    productColl = service.RetrieveMultiple(qProduct);
                } while (productColl.MoreRecords);
                Console.WriteLine("Count !!! " + count);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        private static void createUpdatePriceListItem(Entity pro, IOrganizationService _service)
        {
            try
            {
                done++;

                if (pro.Contains("hil_amount"))
                {
                    Guid DivisionId = pro.GetAttributeValue<EntityReference>("hil_division").Id;
                    Guid ProductId = pro.Id;
                    if (ProductId != Guid.Empty && DivisionId != Guid.Empty && pro.Contains("hil_amount"))
                    {
                        decimal pamount = pro.GetAttributeValue<Money>("hil_amount").Value;
                        // //Console.WriteLine("Amount  " + pamount);
                        Guid iPriceList = Guid.Empty;
                        Guid Uom = Guid.Empty;
                        QueryExpression Query = new QueryExpression("hil_enquirydepartment");
                        Query.ColumnSet = new ColumnSet("hil_pricelist", "hil_uom", "hil_conversionfactor");
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition("hil_productdivision", ConditionOperator.Equal, DivisionId);
                        EntityCollection Found = _service.RetrieveMultiple(Query);
                        if (Found.Entities.Count == 1)
                        {
                            Entity pricelistItem = new Entity("productpricelevel");
                            int conversionFactor = Found.Entities[0].GetAttributeValue<int>("hil_conversionfactor");
                            Decimal caMoney = (pro.GetAttributeValue<Money>("hil_amount").Value) / conversionFactor;
                            iPriceList = Found.Entities[0].GetAttributeValue<EntityReference>("hil_pricelist").Id;
                            Uom = Found.Entities[0].GetAttributeValue<EntityReference>("hil_uom").Id;

                            // //Console.WriteLine("caMoney  " + caMoney);
                            pricelistItem["amount"] = new Money(caMoney);
                            pricelistItem["pricelevelid"] = new EntityReference("pricelevel", iPriceList);
                            pricelistItem["uomid"] = new EntityReference("uom", Uom);
                            pricelistItem["quantitysellingcode"] = new OptionSetValue(1);
                            pricelistItem["productid"] = new EntityReference("product", ProductId);

                            QueryExpression query = new QueryExpression("productpricelevel");
                            query.ColumnSet = new ColumnSet(false);
                            query.Distinct = true;
                            query.Criteria.AddCondition("productid", ConditionOperator.Equal, ProductId);
                            query.Criteria.AddCondition("pricelevelid", ConditionOperator.Equal, iPriceList);
                            EntityCollection productColl = _service.RetrieveMultiple(query);
                            if (productColl.Entities.Count == 1)
                            {
                                if (pamount > 0)
                                {
                                    pricelistItem.Id = productColl.Entities[0].Id;
                                    _service.Update(pricelistItem);
                                    Console.WriteLine("Price List Item is Update " + done + "/" + totlal);

                                }
                                else
                                {
                                    _service.Delete("productpricelevel", productColl.Entities[0].Id);
                                    Console.WriteLine("Price List Item is Deleted" + done + "/" + totlal);
                                }
                            }
                            else
                            {
                                _service.Create(pricelistItem);
                                Console.WriteLine("Price List Item is Created" + done + "/" + totlal);
                            }

                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception " + ex.Message);
            }
        }
    }
}
