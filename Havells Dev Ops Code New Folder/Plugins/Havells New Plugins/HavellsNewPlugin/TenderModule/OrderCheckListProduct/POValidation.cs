using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;


namespace HavellsNewPlugin.TenderModule.OrderCheckListProduct
{
    public class POValidation : IPlugin
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
                    Entity entity = (Entity)context.InputParameters["Target"];
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                    OrderCheckListProductPreCreate.POValidation(service, entity, tracingService);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error " + ex.Message);
            }
        }
    }
}