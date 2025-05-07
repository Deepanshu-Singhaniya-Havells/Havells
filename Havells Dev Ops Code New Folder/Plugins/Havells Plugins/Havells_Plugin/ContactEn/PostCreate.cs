using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.ContactEn
{
    public class PostCreate : IPlugin
    {
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == Contact.EntityLogicalName && context.MessageName.ToUpper() == "CREATE" && context.Depth == 1)
                {
                    tracingService.Trace("1");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Common.CallAPIInsertUpdateCustomer(service, entity, tracingService);
                    //ContactEn.Common.SetFullAddress(entity, context.MessageName.ToUpper(),service);
                }
            }
            catch (Exception ex)
            {
                Entity _address = new Entity(((Entity)context.InputParameters["Target"]).LogicalName);
                _address["hil_errormsg"] = ex.Message;
                _address.Id = ((Entity)context.InputParameters["Target"]).Id;
                service.Update(_address);
            }
        }
    }
}
