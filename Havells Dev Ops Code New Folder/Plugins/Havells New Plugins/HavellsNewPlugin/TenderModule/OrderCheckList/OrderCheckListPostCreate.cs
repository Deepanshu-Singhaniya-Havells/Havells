using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.OrderCheckList
{
    public class OrderCheckListPostCreate : IPlugin
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    tracingService.Trace("1");
                    //Entity entity = (Entity)context.InputParameters["Target"];
                    //if (entity.Contains("hil_tenderno"))
                    //{
                    //    Guid tenderID = entity.GetAttributeValue<EntityReference>("hil_tenderno").Id;
                    //    QueryExpression query1 = new QueryExpression("hil_orderchecklist");
                    //    query1.ColumnSet = new ColumnSet(false);
                    //    query1.Criteria = new FilterExpression(LogicalOperator.And);
                    //    query1.Criteria.AddCondition(new ConditionExpression("hil_tenderno", ConditionOperator.Equal, tenderID));
                    //    EntityCollection entColl1 = service.RetrieveMultiple(query1);
                    //    if (entColl1.Entities.Count != 1)
                    //    {
                    //        throw new InvalidPluginExecutionException("Multiple Order check list for tender is not allowed.  " + entColl1.Entities.Count + "  " + tenderID);
                    //    }
                    //    Entity _tend = new Entity("hil_tender");
                    //    _tend.Id = tenderID;
                    //    _tend["hil_orderchecklist"] = entity.ToEntityReference();
                    //    service.Update(_tend);
                    //    tracingService.Trace("2");
                    //    QueryExpression query = new QueryExpression("hil_tenderproduct");
                    //    query.ColumnSet = new ColumnSet("hil_quantity", "hil_basicpriceinrsmtr");
                    //    query.Criteria = new FilterExpression(LogicalOperator.And);
                    //    query.Criteria.AddCondition(new ConditionExpression("hil_tenderid", ConditionOperator.Equal, tenderID));
                    //    query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                    //    EntityCollection entColl = service.RetrieveMultiple(query);
                    //    foreach (Entity product in entColl.Entities)
                    //    {
                    //        Entity _prd = new Entity(product.LogicalName);
                    //        _prd.Id = product.Id;
                    //        _prd["hil_orderchecklistid"] = entity.ToEntityReference();
                    //        _prd["hil_poqty"] = product.Contains("hil_quantity") ? product.GetAttributeValue<Decimal>("hil_quantity") : 0;
                    //        _prd["hil_porate"] = product.GetAttributeValue<Money>("hil_basicpriceinrsmtr");
                    //        _prd["hil_plantcode"] = entity.GetAttributeValue<EntityReference>("hil_despatchpoint");
                    //        bool inspection = entity.Contains("hil_inspection") ? entity.GetAttributeValue<bool>("hil_inspection") : false;
                    //        if (inspection == true)
                    //        {
                    //            QueryExpression QueryInspectionType = new QueryExpression("hil_inspectiontype");
                    //            QueryInspectionType.ColumnSet = new ColumnSet("hil_name");
                    //            QueryInspectionType.NoLock = true;
                    //            QueryInspectionType.Criteria = new FilterExpression(LogicalOperator.And);
                    //            QueryInspectionType.Criteria.AddCondition("hil_code", ConditionOperator.Equal, "01");
                    //            EntityCollection entCol2 = service.RetrieveMultiple(QueryInspectionType);
                    //            if (entCol2.Entities.Count > 0)
                    //            {
                    //                _prd["hil_inspectiontype"] = entCol2[0].ToEntityReference();
                    //            }
                    //        }
                    //        else
                    //        {
                    //            QueryExpression QueryInspectionType = new QueryExpression("hil_inspectiontype");
                    //            QueryInspectionType.ColumnSet = new ColumnSet("hil_name");
                    //            QueryInspectionType.NoLock = true;
                    //            QueryInspectionType.Criteria = new FilterExpression(LogicalOperator.And);
                    //            QueryInspectionType.Criteria.AddCondition("hil_code", ConditionOperator.Equal, "02");
                    //            EntityCollection entCol2 = service.RetrieveMultiple(QueryInspectionType);
                    //            if (entCol2.Entities.Count > 0)
                    //            {
                    //                _prd["hil_inspectiontype"] = entCol2[0].ToEntityReference();
                    //            }
                    //        }
                    //        service.Update(_prd);
                    //        Entity delSchedule = new Entity("hil_deliveryschedule");
                    //        if (entity.Contains("hil_typeoforder"))
                    //        {
                    //            int typeofOder = entity.GetAttributeValue<OptionSetValue>("hil_typeoforder").Value;
                    //            if (typeofOder == 1)
                    //                delSchedule["hil_deliverydate"] = DateTime.Now.AddDays(7); //DateTime.Now;
                    //            else 
                    //                delSchedule["hil_deliverydate"] = DateTime.Now.AddDays(21);//10 to 21
                    //        }
                    //        delSchedule["hil_quantity"] = product.Contains("hil_quantity") ? product.GetAttributeValue<Decimal>("hil_quantity") : 0; ;
                    //        delSchedule["hil_tenderproduct"] = product.ToEntityReference();
                    //        service.Create(delSchedule);
                    //    }
                    //}
                    //else
                    //{
                    //    int typeofOder = entity.GetAttributeValue<OptionSetValue>("hil_typeoforder").Value;
                    //    if (typeofOder == 1)
                    //    {


                    //    }
                    //    else
                    //    {
                    //        throw new InvalidPluginExecutionException("Tender No. Mandatory...");
                    //    }
                    //}
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.OrderCheckList.OrderCheckListPostCreate.Execute Error " + ex.Message);
            }

        }
    }
}
