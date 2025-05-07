using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.WorkOrder
{
    
    public class PreUpdate_SubStatus : IPlugin
    {
        #region MAIN REGION
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == msdyn_workorder.EntityLogicalName && context.MessageName.ToUpper() == "UPDATE" && context.Depth <= 2)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    msdyn_workorder enWorkorder = entity.ToEntity<msdyn_workorder>();
                    Entity enPreEntity = (Entity)context.PreEntityImages["WorkOrderStatusChanePreEntity"];
                    msdyn_workorder enPreWorkOrder = enPreEntity.ToEntity<msdyn_workorder>();
                    tracingService.Trace("1");
                    if (enWorkorder.msdyn_SubStatus != null && enPreWorkOrder.msdyn_SubStatus != null)
                    {
                        tracingService.Trace("2");
                        StatusTransitionValidation(service, enWorkorder.msdyn_SubStatus, enPreWorkOrder.msdyn_SubStatus);
                    }
                    else if (enWorkorder.msdyn_SubStatus != null && enPreWorkOrder.msdyn_SubStatus == null)
                    {
                        tracingService.Trace("3");
                        StatusTransitionValidation(service, enWorkorder.msdyn_SubStatus, new EntityReference());
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - STATUS CAN'T BE CHANGED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                    }
                }
            }
            catch(Exception ex)
            {
                throw new Exception("ERROR OCCURED AT JOBS.STATUS_UPDATE.PRE-OPERATION.EXECUTE : " + ex.Message.ToUpper());
            }
        }
        #endregion
        #region ON JOB STATUS PRE UPDATE
        public static string GetSubStatucName(IOrganizationService service, Guid statusGuid)
        {
            string _subStatusName = string.Empty;
            QueryExpression qeObj = new QueryExpression("msdyn_workordersubstatus");
            qeObj.ColumnSet = new ColumnSet("msdyn_name");
            qeObj.Criteria = new FilterExpression(LogicalOperator.And);
            qeObj.Criteria.AddCondition(new ConditionExpression("msdyn_workordersubstatusid", ConditionOperator.Equal, statusGuid));
            EntityCollection ecObj = service.RetrieveMultiple(qeObj);
            if (ecObj.Entities.Count == 1)
            {
                _subStatusName = ecObj.Entities[0].GetAttributeValue<string>("msdyn_name");
            }
            return _subStatusName;
        }
        public static void StatusTransitionValidation(IOrganizationService service, EntityReference erNewStatus, EntityReference erBaseStatus)
        {
            try
            {
                QueryExpression iQuery = new QueryExpression("hil_statustransitionmatrix");
                iQuery.ColumnSet = new ColumnSet(false);
                //iQuery.Criteria = new FilterExpression(LogicalOperator.And);
                //iQuery.Criteria.AddCondition("hil_basestatus", ConditionOperator.Equal, erBaseStatus.Id);
                FilterExpression filter = new FilterExpression(LogicalOperator.And);
                FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
                filter1.Conditions.Add(new ConditionExpression("hil_basestatus", ConditionOperator.Equal, erBaseStatus.Id));
                FilterExpression filter2 = new FilterExpression(LogicalOperator.Or);
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition1", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition2", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition3", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition4", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition5", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition6", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition7", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition8", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition9", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition10", ConditionOperator.Equal, erNewStatus.Id));
                filter.AddFilter(filter1);
                filter.AddFilter(filter2);
                iQuery.Criteria = filter;
                EntityCollection ecCollection = service.RetrieveMultiple(iQuery);
                if (ecCollection.Entities.Count == 0)
                {
                    tracingService.Trace("4");
                    throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - STATUS CAN'T BE CHANGED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                    #region UN-USED CONDITIONS
                    //Entity enTransitionMatrix = (Entity)ecCollection.Entities[0];
                    //if (enTransitionMatrix.Contains("hil_statustransition1") && enTransitionMatrix.Attributes.Contains("hil_statustransition1"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition1");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition2") && enTransitionMatrix.Attributes.Contains("hil_statustransition2"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition2");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition3") && enTransitionMatrix.Attributes.Contains("hil_statustransition3"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition3");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition4") && enTransitionMatrix.Attributes.Contains("hil_statustransition4"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition4");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition5") && enTransitionMatrix.Attributes.Contains("hil_statustransition5"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition5");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition6") && enTransitionMatrix.Attributes.Contains("hil_statustransition6"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition6");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition7") && enTransitionMatrix.Attributes.Contains("hil_statustransition7"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition7");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition8") && enTransitionMatrix.Attributes.Contains("hil_statustransition8"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition8");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition9") && enTransitionMatrix.Attributes.Contains("hil_statustransition9"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition9");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition10") && enTransitionMatrix.Attributes.Contains("hil_statustransition10"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition10");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //if (!IfMatch)
                    //{
                    //    throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - STATUS CAN'T BE CHANGED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                    //}
                    #endregion
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ERROR OCCURED AT JOBS.STATUS_UPDATE.PRE-OPERATION.STATUS_TRANSITION_VALIDATION : " + ex.Message.ToUpper());
            }
        }
        #endregion
    }
}