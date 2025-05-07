using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.OrderCheckList
{
   public class OrderCheckListPostUpdate : IPlugin
    {
        public static ITracingService tracingService = null;
        Guid OrderCheckListID;
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
                    Entity entity = (Entity)context.InputParameters["Target"];
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));

                    OrderCheckListID = entity.Id;
                    tracingService.Trace("2");
                    QueryExpression query = new QueryExpression("hil_orderchecklistproduct");
                    query.ColumnSet = new ColumnSet("hil_inspectiontype", "hil_plantcode");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistid", ConditionOperator.Equal, OrderCheckListID));
                    query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                    EntityCollection entColl = service.RetrieveMultiple(query);
                    foreach (Entity product in entColl.Entities)
                    {
                        Entity _prd = new Entity(product.LogicalName);
                        _prd.Id = product.Id;
                        bool inspection = entity.Contains("hil_inspection") ? entity.GetAttributeValue<bool>("hil_inspection") : false;
                        if (inspection == true)
                        {
                            tracingService.Trace("True");
                            QueryExpression QueryInspectionType = new QueryExpression("hil_inspectiontype");
                            QueryInspectionType.ColumnSet = new ColumnSet("hil_name");
                            QueryInspectionType.NoLock = true;
                            QueryInspectionType.Criteria = new FilterExpression(LogicalOperator.And);
                            QueryInspectionType.Criteria.AddCondition("hil_code", ConditionOperator.Equal, "01");
                            EntityCollection entCol2 = service.RetrieveMultiple(QueryInspectionType);
                            if (entCol2.Entities.Count > 0)
                            {
                                _prd["hil_inspectiontype"] = entCol2[0].ToEntityReference();
                            }
                        }
                        else
                        {
                            tracingService.Trace("false");
                            QueryExpression QueryInspectionType = new QueryExpression("hil_inspectiontype");
                            QueryInspectionType.ColumnSet = new ColumnSet("hil_name");
                            QueryInspectionType.NoLock = true;
                            QueryInspectionType.Criteria = new FilterExpression(LogicalOperator.And);
                            QueryInspectionType.Criteria.AddCondition("hil_code", ConditionOperator.Equal, "02");
                            EntityCollection entCol2 = service.RetrieveMultiple(QueryInspectionType);
                            if (entCol2.Entities.Count > 0)
                            {
                                _prd["hil_inspectiontype"] = entCol2[0].ToEntityReference();
                            }
                        }
                        if (entity.Contains("hil_despatchpoint"))
                        {
                            EntityReference plantid = entity.GetAttributeValue<EntityReference>("hil_despatchpoint");
                            _prd["hil_plantcode"] = plantid;
                        }
                        if (entity.Contains("hil_overall"))
                        {
                            decimal tolOverall = entity.GetAttributeValue<decimal>("hil_overall");
                            _prd["hil_tolerancelowerlimit"] = tolOverall;
                            _prd["hil_toleranceupperlimit"] = tolOverall;
                        }
                        service.Update(_prd);
                        tracingService.Trace("Complete");
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.OrderCheckList.OrderCheckListPostUpdate.Execute Error " + ex.Message);
            }

        }
    }
}