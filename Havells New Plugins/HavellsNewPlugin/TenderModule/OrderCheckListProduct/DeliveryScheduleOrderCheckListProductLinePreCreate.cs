using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.OrderCheckListProduct
{
    public class DeliveryScheduleOrderCheckListProductLinePreCreate : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                //if (context.InputParameters.Contains("Target")
                //    && context.InputParameters["Target"] is Entity
                //    && context.PrimaryEntityName.ToLower() == "hil_deliveryschedule"
                //    && context.MessageName.ToUpper() == "CREATE")
                //{
                //    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                //    Entity entity = (Entity)context.InputParameters["Target"];
                //    if (entity.Contains("hil_orderchecklistproduct"))
                //    {
                //        tracingService.Trace("1");
                //        QueryExpression query = new QueryExpression("hil_deliveryschedule");
                //        query.ColumnSet = new ColumnSet("hil_name");
                //        query.Criteria = new FilterExpression(LogicalOperator.And);
                //        query.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistproduct", ConditionOperator.Equal,
                //            entity.GetAttributeValue<EntityReference>("hil_orderchecklistproduct").Id));
                //        query.Orders.Add(new OrderExpression("createdon", OrderType.Descending));
                //        query.TopCount = 1;
                //        EntityCollection entColl = service.RetrieveMultiple(query);
                //        tracingService.Trace("2");
                //        tracingService.Trace("entity.GetAttributeValue<EntityReference>(hil_orderchecklistproduct).Name  " +
                //            entity.GetAttributeValue<EntityReference>("hil_orderchecklistproduct").Name);
                //        int count = 1;
                //        if (entColl.Entities.Count > 0)
                //        {
                //            string _lastTend = entColl[0].GetAttributeValue<string>("hil_name");
                //            string[] number = _lastTend.Split('_');
                //            count = int.Parse(number[1]);
                //            count++;
                //        }
                //        string _tend = service.Retrieve("hil_orderchecklistproduct", entity.GetAttributeValue<EntityReference>("hil_orderchecklistproduct").Id,
                //            new ColumnSet("hil_name")).GetAttributeValue<string>("hil_name");
                //        entity["hil_name"] = _tend + "_" + count.ToString().PadLeft(3, '0');
                //        tracingService.Trace("4");
                //    }
                   
                //}
                //else
                //{
                //    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                //    Entity entity = (Entity)context.InputParameters["Target"];
                //    if (entity.Contains("hil_orderchecklistproduct"))
                //    {

                //        string _tenderProductfetch = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                //                                  <entity name='hil_deliveryschedule'>
                //                                    <attribute name='hil_quantity' alias='amount' aggregate='sum' />
                //                                    <filter type='and'>
                //                                      <condition attribute='hil_orderchecklistproduct' operator='eq' value='" + entity.GetAttributeValue<EntityReference>("hil_orderchecklistproduct").Id + @"' />
                //                                      <condition attribute='statecode' operator='eq' value='0' />
                //                                    </filter>
                //                                  </entity>
                //                                </fetch>";
                //        EntityCollection _tenderproductColl = service.RetrieveMultiple(new FetchExpression(_tenderProductfetch));
                //        tracingService.Trace("_orderchecklistproduct.count" + _tenderproductColl.Entities.Count);
                //        if (_tenderproductColl.Entities.Count > 0)
                //        {
                //            tracingService.Trace("_tenderproductColl");
                //            if (_tenderproductColl[0].Contains("amount"))
                //            {
                //                Money FinalAmount = ((Money)((AliasedValue)_tenderproductColl[0]["amount"]).Value);

                //                QueryExpression query = new QueryExpression("hil_orderchecklistproduct");
                //                query.ColumnSet = new ColumnSet("hil_quantity");
                //                query.Criteria = new FilterExpression(LogicalOperator.And);
                //                query.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistproductid", ConditionOperator.Equal, entity.GetAttributeValue<EntityReference>("hil_orderchecklistproduct").Id));
                //                LinkEntity EntityA = new LinkEntity("hil_orderchecklistproduct", "hil_deliveryschedule", "hil_orderchecklistproductid", "hil_orderchecklistproduct", JoinOperator.Inner);
                //                EntityA.Columns = new ColumnSet("hil_quantity");
                //                EntityA.EntityAlias = "PEnq";
                //                query.LinkEntities.Add(EntityA);
                //                EntityCollection entCol = service.RetrieveMultiple(query);
                //                Decimal poqty = Convert.ToDecimal(entCol[0].GetAttributeValue<decimal>("hil_quantity"));
                //                decimal delQuantity = (Decimal)((AliasedValue)_tenderproductColl[0]["amount"]).Value;
                //                if (poqty < delQuantity)
                //                {
                //                    throw new InvalidPluginExecutionException("Delivery Schedule Quantity Should not be Greater than PO Quanity...");
                //                }

                //            }
                //        }
                //    }
                //}
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.OrderCheckListProduct.DeliveryScheduleOrderCheckListProductLinePreCreate.Execute Error " + ex.Message);
            }
        }

    }
}
