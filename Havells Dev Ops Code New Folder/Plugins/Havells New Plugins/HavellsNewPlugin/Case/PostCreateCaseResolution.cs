using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;

namespace HavellsNewPlugin.Case
{
    public class PostCreateCaseResolution : IPlugin
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
                  && context.PrimaryEntityName.ToLower() == "incidentresolution" && context.Depth == 1 && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("incidentid"))
                    {
                        EntityReference _regardingEntity = entity.GetAttributeValue<EntityReference>("incidentid");
                        string _resolution = entity.GetAttributeValue<string>("subject");
                        _resolution += "\n" + (entity.Contains("description") ? entity.GetAttributeValue<string>("description") : null);
                        Entity _case = new Entity(_regardingEntity.LogicalName, _regardingEntity.Id);
                        _case["adx_resolution"] = _resolution;
                        service.Update(_case);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
        }
    }
}
