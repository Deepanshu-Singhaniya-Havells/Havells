using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.Case
{
    public class CasePostCreate : IPlugin
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                  && context.PrimaryEntityName.ToLower() == "incident" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    //tracingService.Trace(entity.LogicalName);
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_jobid"));
                    EntityReference workOrderRef = entity.GetAttributeValue<EntityReference>("hil_jobid");
                    Entity workOrder = new Entity(workOrderRef.LogicalName, workOrderRef.Id);
                    workOrder["msdyn_servicerequest"] = entity.ToEntityReference();
                    service.Update(workOrder); 
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Case Post Create Error : " + ex.Message);
            }
        }
    }
}
