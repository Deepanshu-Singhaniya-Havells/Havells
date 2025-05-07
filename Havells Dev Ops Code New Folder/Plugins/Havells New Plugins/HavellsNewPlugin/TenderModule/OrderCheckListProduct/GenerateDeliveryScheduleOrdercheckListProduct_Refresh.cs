using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.OrderCheckListProduct
{
    public class GenerateDeliveryScheduleOrdercheckListProduct_Refresh : IPlugin
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
                //if (context.InputParameters.Contains("OrderCheckListId") && context.InputParameters["OrderCheckListId"] is string )
                //{
                //    tracingService.Trace("1");
                //    var OrderCheckListId = context.InputParameters["OrderCheckListId"].ToString();
                //    tracingService.Trace("2");
                //    QueryExpression queryprod = new QueryExpression("hil_orderchecklistproduct");
                //    queryprod.ColumnSet = new ColumnSet("hil_quantity", "hil_name", "hil_orderchecklistid", "createdon");
                //    queryprod.Criteria = new FilterExpression(LogicalOperator.And);
                //    if (OrderCheckListId.Length > 0)
                //    {
                //        queryprod.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistid", ConditionOperator.Equal, new Guid(OrderCheckListId)));
                //    }
                //    LinkEntity EntityQ = new LinkEntity("hil_orderchecklistproduct", "hil_orderchecklist", "hil_orderchecklistid", "hil_orderchecklistid", JoinOperator.Inner);
                //    EntityQ.Columns = new ColumnSet("hil_typeoforder");
                //    EntityQ.EntityAlias = "LnkTypeOfOrder";
                //    queryprod.LinkEntities.Add(EntityQ);
                //    EntityCollection prodCol = service.RetrieveMultiple(queryprod);
                //    tracingService.Trace("3");
                //    foreach (Entity prodentity in prodCol.Entities)
                //    {
                //        Guid OrderCheckListProductID = prodentity.ToEntityReference().Id;
                //        QueryExpression query = new QueryExpression("hil_deliveryschedule");
                //        query.ColumnSet = new ColumnSet(false);
                //        query.Criteria = new FilterExpression(LogicalOperator.And);
                //        query.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistproduct", ConditionOperator.Equal, prodentity.ToEntityReference().Id));
                //        EntityCollection deliveryCol = service.RetrieveMultiple(query);
                //        if (deliveryCol.Entities.Count == 0)
                //        {
                //            Decimal poqty = Convert.ToDecimal(prodentity.GetAttributeValue<decimal>("hil_quantity"));
                //            Entity delSchedule = new Entity("hil_deliveryschedule");
                //            int typeofOder = ((OptionSetValue)prodentity.GetAttributeValue<AliasedValue>("LnkTypeOfOrder.hil_typeoforder").Value).Value;
                //            DateTime createon = prodentity.GetAttributeValue<DateTime>("createdon").Date;
                //            if (typeofOder == 1)
                //                delSchedule["hil_deliverydate"] = createon.AddDays(7);
                //            else
                //                delSchedule["hil_deliverydate"] = createon.AddDays(21);

                //            delSchedule["hil_quantity"] = poqty;
                //            delSchedule["hil_orderchecklistproduct"] = prodentity.ToEntityReference();
                //            service.Create(delSchedule);
                //        }
                //        else if (deliveryCol.Entities.Count == 1)
                //        {
                //            Decimal poqty = Convert.ToDecimal(prodentity.GetAttributeValue<decimal>("hil_quantity"));
                //            Entity delSchedule = new Entity("hil_deliveryschedule");
                //            delSchedule.Id = deliveryCol[0].Id;
                //            int typeofOder = ((OptionSetValue)prodentity.GetAttributeValue<AliasedValue>("LnkTypeOfOrder.hil_typeoforder").Value).Value;
                //            DateTime createon = prodentity.GetAttributeValue<DateTime>("createdon").Date;
                //            if (typeofOder == 1)
                //                delSchedule["hil_deliverydate"] = createon.AddDays(7);
                //            else
                //                delSchedule["hil_deliverydate"] = createon.AddDays(21);
                //            delSchedule["hil_quantity"] = poqty;
                //            delSchedule["hil_orderchecklistproduct"] = prodentity.ToEntityReference();
                //            service.Update(delSchedule);
                //        }

                //    }
                //    #region autoCorrect Serial No.
                //    QueryExpression qsChecklist = new QueryExpression("hil_orderchecklistproduct");
                //    qsChecklist.ColumnSet = new ColumnSet("hil_name", "hil_orderchecklistid");
                //    qsChecklist.Criteria = new FilterExpression(LogicalOperator.And);
                //    qsChecklist.Criteria.AddCondition("hil_orderchecklistid", ConditionOperator.Equal, new Guid(OrderCheckListId));
                //    EntityCollection entCol = service.RetrieveMultiple(qsChecklist);
                //    int count = entCol.Entities.Count;
                //    throw new InvalidPluginExecutionException(count);
                //    int i = 1;
                //    foreach (Entity prod in entCol.Entities)
                //    {
                //        string _tend = prod.GetAttributeValue<EntityReference>("hil_orderchecklistid").Name;
                //        string name = prod.GetAttributeValue<string>("hil_name");
                //        Entity prd = new Entity(prod.LogicalName);
                //        prd.Id = prod.Id;
                //        prd["hil_name"] = _tend + "_" + i.ToString().PadLeft(3, '0');
                //        service.Update(prd);
                //        i++;
                //        QueryExpression qsChecklist1 = new QueryExpression("hil_deliveryschedule");
                //        qsChecklist1.ColumnSet = new ColumnSet("hil_name", "hil_orderchecklistproduct");
                //        qsChecklist1.Criteria = new FilterExpression(LogicalOperator.And);
                //        qsChecklist1.Criteria.AddCondition("hil_orderchecklistproduct", ConditionOperator.Equal, prod.Id);
                //        EntityCollection entCol1 = service.RetrieveMultiple(qsChecklist1);
                //        foreach (Entity prod1 in entCol1.Entities)
                //        {
                //            string _tndprd = prod1.GetAttributeValue<EntityReference>("hil_orderchecklistproduct").Name;
                //            string name1 = prod1.GetAttributeValue<string>("hil_name");
                //            int j = 1;
                //            Entity prd1 = new Entity(prod1.LogicalName);
                //            prd1.Id = prod1.Id;
                //            prd1["hil_name"] = _tndprd + "_" + j.ToString().PadLeft(3, '0');
                //            j++;
                //            service.Update(prd1);
                //        }
                //    }
                //    #endregion
                //}

            }
            catch (Exception ex)
            {
                throw new Exception("Error :" + ex.Message);
            }
        }
    }
}
