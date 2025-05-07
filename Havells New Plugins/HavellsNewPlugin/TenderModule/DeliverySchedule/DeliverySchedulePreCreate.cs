using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.DeliverySchedule
{
    public class DeliverySchedulePreCreate : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            //#region PluginConfig
            //tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            //IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            //IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            //IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            //#endregion
            //try
            //{
            //    if (context.InputParameters.Contains("Target")
            //        && context.InputParameters["Target"] is Entity
            //        && context.PrimaryEntityName.ToLower() == "hil_deliveryschedule"
            //        && context.MessageName.ToUpper() == "CREATE")
            //    {
            //        OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
            //        Entity entity = (Entity)context.InputParameters["Target"];
            //        if (entity.Contains("hil_tenderproduct"))
            //        {
            //            QueryExpression query = new QueryExpression("hil_deliveryschedule");
            //            query.ColumnSet = new ColumnSet(false);
            //            query.Criteria = new FilterExpression(LogicalOperator.And);
            //            query.Criteria.AddCondition(new ConditionExpression("hil_tenderproduct", ConditionOperator.Equal, entity.GetAttributeValue<EntityReference>("hil_tenderproduct").Id));
            //            EntityCollection entColl = service.RetrieveMultiple(query);
            //            int count = entColl.Entities.Count + 1;
            //            string _tend = service.Retrieve("hil_tenderproduct", entity.GetAttributeValue<EntityReference>("hil_tenderproduct").Id, new ColumnSet("hil_name")).GetAttributeValue<string>("hil_name");
            //            entity["hil_name"] = _tend + "_" + count.ToString().PadLeft(3, '0');
            //        }
            //    }
            //    else
            //    {
            //        OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
            //        Entity entity = (Entity)context.InputParameters["Target"];
            //        if (entity.Contains("hil_tenderproduct"))
            //        {

            //            string _tenderProductfetch = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
            //                                      <entity name='hil_deliveryschedule'>
            //                                        <attribute name='hil_tenderid' alias='tenderNo' groupby='true' />
            //                                        <attribute name='hil_quantity' alias='amount' aggregate='sum' />
            //                                        <filter type='and'>
            //                                          <condition attribute='hil_tenderproduct' operator='eq' value='" + entity.GetAttributeValue<EntityReference>("hil_tenderproduct").Id + @"' />
            //                                          <condition attribute='statecode' operator='eq' value='0' />
            //                                        </filter>
            //                                      </entity>
            //                                    </fetch>";
            //            EntityCollection _tenderproductColl = service.RetrieveMultiple(new FetchExpression(_tenderProductfetch));
            //            tracingService.Trace("_tenderproductColl.count" + _tenderproductColl.Entities.Count);
            //            if (_tenderproductColl.Entities.Count > 0)
            //            {
            //                tracingService.Trace("_tenderproductColl");
            //                if (_tenderproductColl[0].Contains("amount"))
            //                {
            //                    Money FinalAmount = ((Money)((AliasedValue)_tenderproductColl[0]["amount"]).Value);

            //                    QueryExpression query = new QueryExpression("hil_tenderproduct");
            //                    query.ColumnSet = new ColumnSet("hil_poqty");
            //                    query.Criteria = new FilterExpression(LogicalOperator.And);
            //                    query.Criteria.AddCondition(new ConditionExpression("hil_tenderproductid", ConditionOperator.Equal, "7096b386-3a0e-ec11-b6e6-002248d4bfe1"));
            //                    LinkEntity EntityA = new LinkEntity("hil_tenderproduct", "hil_deliveryschedule", "hil_tenderproductid", "hil_tenderproduct", JoinOperator.Inner);
            //                    EntityA.Columns = new ColumnSet("hil_quantity");
            //                    EntityA.EntityAlias = "PEnq";
            //                    query.LinkEntities.Add(EntityA);
            //                    EntityCollection entCol = service.RetrieveMultiple(query);
            //                    Decimal poqty = Convert.ToDecimal(entCol[0].GetAttributeValue<decimal>("hil_poqty"));
            //                    decimal delQuantity = (Decimal)((AliasedValue)_tenderproductColl[0]["amount"]).Value;
            //                    if (poqty < delQuantity)
            //                    {
            //                        throw new InvalidPluginExecutionException("Delivery Schedule Quantity Should not be Greater than PO Quanity...");
            //                    }

            //                }
            //            }
            //        }
                        
            //    }
            //}
            //catch (Exception ex)
            //{
            //    tracingService.Trace("Error " + ex.Message);
            //    throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.DeliverySchedule.Execute Error " + ex.Message);
            //}
        }
    }
}
