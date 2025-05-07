using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.WorkOrder
{
    public class PostUpdate_AutoAssign : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            try
            {
                tracingService.Trace("DEPTH : "+ context.Depth.ToString());
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_workorder" && context.MessageName.ToUpper() == "UPDATE"
                    && context.Depth < 2)
                {
                    //tracingService.Trace("2");
                    Entity preentity = (Entity)context.PreEntityImages["PreImage_Update"];
                    msdyn_workorder PreenWO = preentity.ToEntity<msdyn_workorder>();
                    Entity entity = (Entity)context.InputParameters["Target"];
                    msdyn_workorder Ent = entity.ToEntity<msdyn_workorder>();
                    if(entity.Contains("hil_automaticassign") && Ent.hil_AutomaticAssign != null && Ent.hil_AutomaticAssign.Value == 1)
                    {
                        if((PreenWO.msdyn_SubStatus == null || (PreenWO.msdyn_SubStatus != null && PreenWO.msdyn_SubStatus.Name == "Pending for Allocation") || (PreenWO.msdyn_SubStatus != null && PreenWO.msdyn_SubStatus.Name == "Registered")))
                        {
                            Entity iJob = service.Retrieve(msdyn_workorder.EntityLogicalName, entity.Id, new ColumnSet(true));
                            Havells_Plugin.WorkOrder.PostUpdate_Asynch.CallAllocation(service, entity, tracingService, iJob);
                            Plugins.ServiceTicket.PostCreate.PopulateBrand(iJob, service);
                            Plugins.ServiceTicket.PostCreate.SetKKGCode(iJob, service);
                        }
                    }
                    if (entity.Contains("hil_callsubtype") || entity.Contains("hil_productsubcategory") || entity.Contains("hil_address"))
                    {
                        if (PreenWO.msdyn_SubStatus.Name != "Closed" && PreenWO.msdyn_SubStatus.Name != "Canceled" && PreenWO.msdyn_SubStatus.Name != "Work Done" && PreenWO.msdyn_SubStatus.Name != "Work Done SMS")
                        {
                            Entity iJob = service.Retrieve(msdyn_workorder.EntityLogicalName, entity.Id, new ColumnSet(true));
                            Havells_Plugin.WorkOrder.PostUpdate_Asynch.CallAllocation(service, entity, tracingService, iJob);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrder.PostUpdate_AutoAssign.Execute : " + ex.Message.ToUpper());
            }
            #endregion
        }
    }
}
