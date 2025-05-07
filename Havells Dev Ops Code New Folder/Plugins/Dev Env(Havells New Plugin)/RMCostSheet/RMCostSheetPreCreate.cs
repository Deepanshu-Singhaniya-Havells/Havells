using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace HavellsNewPlugin.RMCostSheet
{
    public class RMCostSheetPreCreate : IPlugin
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
                if (context.MessageName.ToLower() == "create"
                    && context.InputParameters["Target"] is Entity
                    && context.InputParameters.Contains("Target"))
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.Id != Guid.Empty)
                    {
                        tracingService.Trace("GetPackageCharges Method");
                        Guid TenderId = entity.Contains("hil_tenderid") ? entity.GetAttributeValue<EntityReference>("hil_tenderid").Id : Guid.Empty;
                        Entity entityTender = service.Retrieve("hil_tender", TenderId, new ColumnSet("hil_department"));
                        Guid TenderDepartmentId = entityTender.Contains("hil_department") ? entityTender.GetAttributeValue<EntityReference>("hil_department").Id : Guid.Empty;

                        QueryExpression query = new QueryExpression("hil_departmentrmcostsetup");
                        query.ColumnSet = new ColumnSet("hil_mfgvariableexpenseper", "hil_packagingexpenseper", "hil_mfgfixedexpenseper");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition(new ConditionExpression("hil_department", ConditionOperator.Equal, TenderDepartmentId));
                        EntityCollection entColl = service.RetrieveMultiple(query);
                        if (entColl.Entities.Count > 0)
                        {
                            if (entColl.Entities[0].Contains("hil_mfgvariableexpenseper"))
                                entity["hil_mfgvariableexpenseper"] = entColl.Entities[0].GetAttributeValue<decimal>("hil_mfgvariableexpenseper");
                            if (entColl.Entities[0].Contains("hil_packagingexpenseper"))
                                entity["hil_packagingexpenseper"] = entColl.Entities[0].GetAttributeValue<decimal>("hil_packagingexpenseper");
                            if (entColl.Entities[0].Contains("hil_mfgfixedexpenseper"))
                                entity["hil_mfgfixedexpenseper"] = entColl.Entities[0].GetAttributeValue<decimal>("hil_mfgfixedexpenseper");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("HavellsNewPlugin.RMCostSheetLine.RMCostSheetPreCreate.Execute Error " + ex.Message);
            }
        }
    }
}
