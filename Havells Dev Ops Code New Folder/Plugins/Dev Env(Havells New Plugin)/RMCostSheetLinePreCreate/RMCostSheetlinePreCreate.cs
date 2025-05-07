using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.RMCostSheetLine
{
    public class RMCostSheetlinePreCreate : IPlugin
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
                Guid rmcostsheetid = Guid.Empty;
                string fetchXml = "";
                if ((((context.MessageName.ToLower() == "create" || context.MessageName.ToLower() == "update") && context.InputParameters["Target"] is Entity)
                    || (context.MessageName.ToLower() == "delete" && context.InputParameters["Target"] is EntityReference))
                   && context.InputParameters.Contains("Target") && context.Depth == 1)
                {

                    if (context.MessageName.ToLower() == "delete")
                    {
                        tracingService.Trace("Delete");
                        EntityReference entityref = (EntityReference)context.InputParameters["Target"];
                        Entity entSheetline = service.Retrieve(entityref.LogicalName, entityref.Id, new ColumnSet("hil_rmcostsheet"));
                        rmcostsheetid = entSheetline.Contains("hil_rmcostsheet") ? entSheetline.GetAttributeValue<EntityReference>("hil_rmcostsheet").Id : Guid.Empty;

                        fetchXml = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                            <entity name='hil_rmcostsheetline'>
                            <attribute name='hil_rmcostsheet' alias='rmcostsheet' groupby='true' />
                            <attribute name='hil_cost' alias='cost' aggregate='sum' />  
                            <attribute name='hil_rmcostsheet' alias='rmcostsheetCount' aggregate='count' />
                                <filter type='and'>
                                <condition attribute='hil_rmcostsheet' operator='eq' value='" + rmcostsheetid + @"' />
                                <condition attribute='hil_rmcostsheetlineid' operator='ne' value='" + entityref.Id + @"' />
                                <condition attribute='statecode' operator='eq' value='0' />
                                </filter>
                            </entity>
                        </fetch>";
                    }
                    else
                    {
                        tracingService.Trace("create");
                        Entity entity = (Entity)context.InputParameters["Target"];
                        Entity entSheetline = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_rmcostsheet"));
                        rmcostsheetid = entSheetline.Contains("hil_rmcostsheet") ? entSheetline.GetAttributeValue<EntityReference>("hil_rmcostsheet").Id : Guid.Empty;

                        fetchXml = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                            <entity name='hil_rmcostsheetline'>
                            <attribute name='hil_rmcostsheet' alias='rmcostsheet' groupby='true' />
                            <attribute name='hil_cost' alias='cost' aggregate='sum' />  
                            <attribute name='hil_rmcostsheet' alias='rmcostsheetCount' aggregate='count' />
                                <filter type='and'>
                                <condition attribute='hil_rmcostsheet' operator='eq' value='" + rmcostsheetid + @"' />
                                <condition attribute='statecode' operator='eq' value='0' />
                                </filter>
                            </entity>
                        </fetch>";
                    }
                    tracingService.Trace(rmcostsheetid.ToString());
                    if (rmcostsheetid != Guid.Empty)
                    {
                        tracingService.Trace("after create");
                        EntityCollection RMCostSheetlineColl = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (RMCostSheetlineColl.Entities.Count > 0)
                        {

                            Money TotalRMCost = (Money)((AliasedValue)RMCostSheetlineColl[0]["cost"]).Value;
                            tracingService.Trace(TotalRMCost.ToString());
                            tracingService.Trace(TotalRMCost.ToString());
                            int rmcostsheetCount = (int)((AliasedValue)RMCostSheetlineColl[0]["rmcostsheetCount"]).Value;
                            tracingService.Trace(rmcostsheetCount.ToString());
                            Entity RMCostSheet = new Entity("hil_rmcostsheet", rmcostsheetid);
                            RMCostSheet["hil_rmcost"] = TotalRMCost;


                            QueryExpression query = new QueryExpression("hil_departmentrmcostsetup");
                            query.ColumnSet = new ColumnSet("hil_mfgvariableexpenseper", "hil_packagingexpenseper", "hil_mfgfixedexpenseper");
                            EntityCollection entColl = service.RetrieveMultiple(query);
                            if (entColl.Entities.Count > 0)
                            {
                                if (entColl.Entities[0].Contains("hil_mfgvariableexpenseper"))
                                    RMCostSheet["hil_mfgvariableexpenseper"] = entColl.Entities[0].GetAttributeValue<decimal>("hil_mfgvariableexpenseper");
                                if (entColl.Entities[0].Contains("hil_packagingexpenseper"))
                                    RMCostSheet["hil_packagingexpenseper"] = entColl.Entities[0].GetAttributeValue<decimal>("hil_packagingexpenseper");
                                if (entColl.Entities[0].Contains("hil_mfgfixedexpenseper"))
                                    RMCostSheet["hil_mfgfixedexpenseper"] = entColl.Entities[0].GetAttributeValue<decimal>("hil_mfgfixedexpenseper");
                            }

                            service.Update(RMCostSheet);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("HavellsNewPlugin.RMCostSheetLine.RMCostSheetlinePreCreate.Execute Error " + ex.Message);
            }
        }
    }
}