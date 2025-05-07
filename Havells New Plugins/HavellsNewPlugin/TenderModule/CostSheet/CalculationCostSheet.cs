using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.CostSheet
{
    public class CalculationCostSheet : IPlugin
    {
        public static ITracingService tracingService = null;
        //Guid costsheetheader;
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
            //    if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            //    {
            //        tracingService.Trace("1");
            //        Entity entity = (Entity)context.InputParameters["Target"];
            //        tracingService.Trace("2");
            //        costsheetheader = ((EntityReference)service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_costsheetheader"))["hil_costsheetheader"]).Id;
            //        string _tenderProductfetch = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
            //                                      <entity name='"+entity.LogicalName+@"'>   
            //                                        <attribute name='hil_costsheetheader' alias='hil_costsheetheader' groupby='true' />
            //                                        <attribute name='hil_materialcost' alias='amount' aggregate='sum' />
            //                                        <filter type='and'>
            //                                          <condition attribute='hil_costsheetheader' operator='eq' value='" + costsheetheader + @"' />
            //                                          <condition attribute='statecode' operator='eq' value='0' />
            //                                        </filter>
            //                                      </entity>
            //                                    </fetch>";
            //        EntityCollection hil_costsheetheaderLineColl = service.RetrieveMultiple(new FetchExpression(_tenderProductfetch));
            //        tracingService.Trace("_tenderproductColl.count" + hil_costsheetheaderLineColl.Entities.Count);
            //        if (hil_costsheetheaderLineColl.Entities.Count > 0)
            //        {
            //            tracingService.Trace("hil_costsheetheaderLineCollColl");
            //            if (hil_costsheetheaderLineColl[0].Contains("amount"))
            //            {
            //                Money FinalAmount = ((Money)((AliasedValue)hil_costsheetheaderLineColl[0]["amount"]).Value);

            //                tracingService.Trace("hil_costsheetheaderLineCollColl.count" + FinalAmount.Value);
            //                Entity hil_costsheetheader = new Entity("hil_costsheetheader");
            //                hil_costsheetheader["hil_rowmaterialcost"] = FinalAmount;
            //                hil_costsheetheader.Id = costsheetheader;
            //                service.Update(hil_costsheetheader);
            //                tracingService.Trace("Tender Uploaded");
            //            }
            //        }

            //    }
            //}
            //catch (Exception ex)
            //{
            //    tracingService.Trace("Error " + ex.Message);
            //    throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.CostSheet.CalculationCostSheet.Execute Error " + ex.Message);
            //}

        }
    }
}
